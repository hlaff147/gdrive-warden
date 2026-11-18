/**
 * ============================================================================
 * PROJETO A: AUDITORIA E DEDUPLICAÇÃO (SOMENTE LEITURA)
 * ============================================================================
 * Escopo de Segurança: drive.readonly (Incapaz de alterar arquivos no Drive).
 * Gera a planilha 'Relatorio_Auditoria_Drive' com as abas:
 * Inventario | Duplicatas_Detectadas | Ignorados | Resumo | Config | Log
 */

const CONFIG = {
  DRY_RUN: true,
  PASTA_RAIZ_DO_ESCOPO: '', // Opcional: ID de '_TESTE_ORGANIZACAO' para escopo restrito
  PERFIL_ESCOPO: 'CONSERVADOR', // CONSERVADOR | EQUILIBRADO | AMPLO
  TAMANHO_LOTE_PILOTO: 20,
  LIMITE_ERROS_CONSECUTIVOS: 5,
  TEMPO_MAXIMO_MS: 270000, // 4,5 min (margem segura antes dos 6 min do Apps Script)
  FRASE_DE_CONFIRMACAO: 'EU REVISEI A PLANILHA E AUTORIZO A EXECUCAO',
  NOME_PLANILHA_RELATORIO: 'Relatorio_Auditoria_Drive',
  LIMITE_CELULAS_PLANILHA: 9000000,
  TAMANHO_LOTE_ESCRITA: 2500,
  PASTAS_PROTEGIDAS_NOMES: [
    'Colab Notebooks',
    '_QUARENTENA_DUPLICATAS_PARA_REVISAO',
    '_TRIAGEM_MANUAL',
    '_RESTAURADOS',
    '_TESTE_ORGANIZACAO'
  ],
  MARCADORES_DE_PROJETO: [
    '.git', 'package.json', 'pyproject.toml', 'requirements.txt',
    'pom.xml', 'build.gradle', 'go.mod', 'cargo.toml', 'dockerfile',
    'docker-compose.yml', 'makefile', 'tsconfig.json', 'main.tex'
  ],
  PASTAS_ESTRUTURA_PROJETO: [
    'src', 'services', 'k6', 'analysis', 'artigo_latex'
  ],
  PASTAS_GENERICAS_POUCO_SIGNIFICATIVAS: [
    'meu drive', 'root', 'novos uploads', 'downloads', 'temp', 'triagem', 'sem título', 'untitled', 'camera'
  ]
};

function etapa1_auditoria() {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    Logger.log('Execução abortada: já existe outra instância de etapa1_auditoria em andamento.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  let ss = null;

  try {
    ss = obterOuCriarPlanilhaRelatorio_(props);
    registrarLog_(ss, 'etapa1_auditoria', 'SISTEMA', 'INICIO_OU_RETOMADA', 'Escopo: ' + (CONFIG.PASTA_RAIZ_DO_ESCOPO || 'DRIVE_INTEIRO'));

    let fase = props.getProperty('AUDITORIA_FASE') || 'COLETA';

    if (fase === 'COLETA') {
      const concluiuColeta = executarFaseColeta_(ss, props, startTime);
      if (!concluiuColeta) {
        agendarGatilhoContinuacao_('etapa1_auditoria');
        registrarLog_(ss, 'etapa1_auditoria', 'SISTEMA', 'PAUSA_TEMPO_COLETA', 'Gatilho agendado para continuar coleta.');
        return;
      }
      props.setProperty('AUDITORIA_FASE', 'PROCESSAMENTO');
      fase = 'PROCESSAMENTO';
    }

    if (fase === 'PROCESSAMENTO') {
      if (Date.now() - startTime > CONFIG.TEMPO_MAXIMO_MS - 45000) {
        agendarGatilhoContinuacao_('etapa1_auditoria');
        registrarLog_(ss, 'etapa1_auditoria', 'SISTEMA', 'PAUSA_ANTES_PROCESSAMENTO', 'Gatilho agendado para fase de processamento.');
        return;
      }

      executarFaseProcessamento_(ss);
      limparEstadoEGatilhos_(props);
      registrarLog_(ss, 'etapa1_auditoria', 'SISTEMA', 'CONCLUIDO', 'Auditoria finalizada com sucesso. URL: ' + ss.getUrl());
      Logger.log('Auditoria concluída! Planilha privada disponível em: ' + ss.getUrl());
    }
  } catch (erro) {
    limparGatilhosContinuacao_();
    if (ss) {
      registrarLog_(ss, 'etapa1_auditoria', 'ERRO_FATAL', 'FALHA', erro.message + ' | Stack: ' + (erro.stack || ''));
    }
    throw erro;
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function executarFaseColeta_(ss, props, startTime) {
  const abaInv = ss.getSheetByName('Inventario');
  let pageToken = props.getProperty('CURSOR_PAGE_TOKEN') || null;
  let totalColetado = Number(props.getProperty('TOTAL_COLETADO') || 0);
  let bufferLinhas = [];

  const fields = [
    'nextPageToken',
    'files(id,name,mimeType,md5Checksum,sha256Checksum,size,createdTime,modifiedTime,parents,ownedByMe,shared,starred,description,trashed,driveId,capabilities(canMoveItemWithinDrive),shortcutDetails,imageMediaMetadata(time,width,height,cameraMake,cameraModel))'
  ].join(',');

  do {
    if (Date.now() - startTime > CONFIG.TEMPO_MAXIMO_MS) {
      if (bufferLinhas.length > 0) {
        despejarBufferNaAba_(abaInv, bufferLinhas);
        totalColetado += bufferLinhas.length;
        bufferLinhas = [];
      }
      props.setProperty('CURSOR_PAGE_TOKEN', pageToken);
      props.setProperty('TOTAL_COLETADO', String(totalColetado));
      return false;
    }

    const params = {
      q: "'me' in owners and trashed = false",
      pageSize: 1000,
      pageToken: pageToken || undefined,
      fields: fields,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    };

    const resposta = chamarApiComRetry_(() => Drive.Files.list(params));
    const arquivos = resposta.files || [];

    for (let i = 0; i < arquivos.length; i++) {
      const f = arquivos[i];
      const parentId = (f.parents && f.parents.length > 0) ? f.parents[0] : 'root';
      const canMove = (f.capabilities && typeof f.capabilities.canMoveItemWithinDrive === 'boolean')
        ? f.capabilities.canMoveItemWithinDrive
        : false;
      const isShortcut = Boolean(f.mimeType === 'application/vnd.google-apps.shortcut' || f.shortcutDetails);
      const exif = f.imageMediaMetadata || {};
      const dimensoes = (exif.width && exif.height) ? (exif.width + 'x' + exif.height) : '';
      const camera = [exif.cameraMake || '', exif.cameraModel || ''].join(' ').trim();

      bufferLinhas.push([
        f.id || '',
        f.name || '',
        f.mimeType || '',
        '',
        parentId,
        f.md5Checksum || '',
        f.sha256Checksum || '',
        f.size ? String(f.size) : '',
        f.createdTime || '',
        f.modifiedTime || '',
        exif.time || '',
        dimensoes,
        camera,
        f.ownedByMe ? 'SIM' : 'NAO',
        f.shared ? 'SIM' : 'NAO',
        f.driveId || '',
        canMove ? 'SIM' : 'NAO',
        isShortcut ? 'SIM' : 'NAO',
        (f.starred || f.description) ? 'SIM' : 'NAO',
        '',
        ''
      ]);

      if (bufferLinhas.length >= CONFIG.TAMANHO_LOTE_ESCRITA) {
        despejarBufferNaAba_(abaInv, bufferLinhas);
        totalColetado += bufferLinhas.length;
        bufferLinhas = [];
      }
    }

    pageToken = resposta.nextPageToken || null;
  } while (pageToken);

  if (bufferLinhas.length > 0) {
    despejarBufferNaAba_(abaInv, bufferLinhas);
    totalColetado += bufferLinhas.length;
  }

  props.deleteProperty('CURSOR_PAGE_TOKEN');
  props.setProperty('TOTAL_COLETADO', String(totalColetado));
  return true;
}

function executarFaseProcessamento_(ss) {
  const abaInv = ss.getSheetByName('Inventario');
  const abaDup = ss.getSheetByName('Duplicatas_Detectadas');
  const abaIgn = ss.getSheetByName('Ignorados');
  const abaRes = ss.getSheetByName('Resumo');

  const ultimaLinha = abaInv.getLastRow();
  if (ultimaLinha <= 1) {
    escreverResumo_(abaRes, [], [], { linhasPlanilha: [], bytesEconomiaPotencialN1: 0 });
    return;
  }

  const dadosBrutos = abaInv.getRange(2, 1, ultimaLinha - 1, 21).getValues();
  const itens = dadosBrutos.map((row, idx) => ({
    rowIndex: idx,
    id: String(row[0]),
    name: String(row[1]),
    mimeType: String(row[2]),
    path: '',
    parentId: String(row[4] || 'root'),
    md5Checksum: String(row[5] || ''),
    sha256Checksum: String(row[6] || ''),
    size: row[7] !== '' ? Number(row[7]) : null,
    createdTime: String(row[8] || ''),
    modifiedTime: String(row[9] || ''),
    exifTime: String(row[10] || ''),
    dimensions: String(row[11] || ''),
    camera: String(row[12] || ''),
    ownedByMe: row[13] === 'SIM',
    shared: row[14] === 'SIM',
    driveId: String(row[15] || ''),
    canMove: row[16] === 'SIM',
    isShortcut: row[17] === 'SIM',
    hasStarOrDesc: row[18] === 'SIM',
    isProtected: false,
    sharedInherited: false,
    inScope: true,
    skipReason: ''
  }));

  const contextoArvore = construirContextoDeArvorePuro_(itens, CONFIG);

  const ignoradosRows = [];
  for (let i = 0; i < itens.length; i++) {
    const it = itens[i];
    dadosBrutos[i][3] = it.path;
    dadosBrutos[i][19] = it.isProtected ? 'SIM' : 'NAO';
    dadosBrutos[i][20] = it.skipReason ? ('IGNORADO: ' + it.skipReason) : 'ELEGIVEL';

    if (it.skipReason && it.inScope) {
      ignoradosRows.push([it.id, it.name, it.mimeType, it.path, it.skipReason]);
    }
  }

  regravaMatrizEmLotes_(abaInv, 2, 1, dadosBrutos);

  const resultadoDedup = executarMotorDeduplicacaoPuro_(itens, contextoArvore.pastasOrganizadasIds, CONFIG);

  limparDadosMantendoCabecalho_(abaDup);
  if (resultadoDedup.linhasPlanilha.length > 0) {
    regravaMatrizEmLotes_(abaDup, 2, 1, resultadoDedup.linhasPlanilha);
  }

  limparDadosMantendoCabecalho_(abaIgn);
  if (ignoradosRows.length > 0) {
    regravaMatrizEmLotes_(abaIgn, 2, 1, ignoradosRows);
  }

  escreverResumo_(abaRes, itens, ignoradosRows, resultadoDedup);
}

function construirContextoDeArvorePuro_(itens, config) {
  const folderMap = {};
  const nomesProtegidosSet = new Set((config.PASTAS_PROTEGIDAS_NOMES || []).map(n => n.toLowerCase()));
  const marcadoresProjetoSet = new Set((config.MARCADORES_DE_PROJETO || []).map(n => n.toLowerCase()));
  const pastasEstruturaSet = new Set((config.PASTAS_ESTRUTURA_PROJETO || []).map(n => n.toLowerCase()));
  const pastasGenericasSet = new Set((config.PASTAS_GENERICAS_POUCO_SIGNIFICATIVAS || []).map(n => n.toLowerCase()));

  for (let i = 0; i < itens.length; i++) {
    const it = itens[i];
    if (it.mimeType === 'application/vnd.google-apps.folder') {
      folderMap[it.id] = {
        id: it.id,
        name: it.name,
        parentId: it.parentId || 'root',
        shared: Boolean(it.shared || !it.ownedByMe || it.driveId),
        protected: nomesProtegidosSet.has(it.name.toLowerCase())
      };
    }
  }

  for (let i = 0; i < itens.length; i++) {
    const it = itens[i];
    const nomeLower = it.name.toLowerCase();
    const isSln = nomeLower.endsWith('.sln');
    const isMarkerFile = marcadoresProjetoSet.has(nomeLower) || isSln;
    const isProjectSubdir = (it.mimeType === 'application/vnd.google-apps.folder') && pastasEstruturaSet.has(nomeLower);

    if ((isMarkerFile || isProjectSubdir) && it.parentId && folderMap[it.parentId]) {
      const raizProjetoId = obterAncestralDeTopoDaPasta_(it.parentId, folderMap, config.PASTA_RAIZ_DO_ESCOPO);
      if (folderMap[raizProjetoId]) {
        folderMap[raizProjetoId].protected = true;
      }
      folderMap[it.parentId].protected = true;
    }
  }

  let mudou = true;
  while (mudou) {
    mudou = false;
    const folderIds = Object.keys(folderMap);
    for (let i = 0; i < folderIds.length; i++) {
      const f = folderMap[folderIds[i]];
      const pai = folderMap[f.parentId];
      if (pai) {
        if (pai.protected && !f.protected) {
          f.protected = true;
          mudou = true;
        }
        if (pai.shared && !f.shared) {
          f.shared = true;
          mudou = true;
        }
      }
    }
  }

  const pastasOrganizadasIds = new Set();
  Object.keys(folderMap).forEach(fid => {
    const f = folderMap[fid];
    if (!pastasGenericasSet.has(f.name.trim().toLowerCase())) {
      pastasOrganizadasIds.add(fid);
    }
  });

  const escopoRaizId = (config.PASTA_RAIZ_DO_ESCOPO || '').trim();

  for (let i = 0; i < itens.length; i++) {
    const it = itens[i];
    it.path = resolverCaminhoPuro_(it.parentId, folderMap);

    if (escopoRaizId) {
      it.inScope = pertenceSubarvore_(it.id, it.parentId, escopoRaizId, folderMap);
    } else {
      it.inScope = true;
    }

    const pastaPai = folderMap[it.parentId];
    it.isProtected = Boolean(
      (it.mimeType === 'application/vnd.google-apps.folder' && folderMap[it.id] && folderMap[it.id].protected) ||
      (pastaPai && pastaPai.protected) ||
      nomesProtegidosSet.has(it.name.toLowerCase())
    );
    it.sharedInherited = Boolean(it.shared || (pastaPai && pastaPai.shared));

    if (!it.inScope) {
      it.skipReason = 'FORA_DA_PASTA_RAIZ_DO_ESCOPO';
    } else if (!it.ownedByMe) {
      it.skipReason = 'NAO_PROPRIETARIO (ownedByMe=false)';
    } else if (it.driveId) {
      it.skipReason = 'DRIVE_COMPARTILHADO (driveId presente)';
    } else if (it.isShortcut) {
      it.skipReason = 'ATALHO (shortcut)';
    } else if (it.mimeType === 'application/vnd.google-apps.folder' && it.sharedInherited) {
      it.skipReason = 'PASTA_COMPARTILHADA';
    } else if (it.sharedInherited) {
      it.skipReason = 'ARQUIVO_OU_PASTA_PAI_COMPARTILHADO';
    } else if (!it.canMove) {
      it.skipReason = 'SEM_PERMISSAO_MOVER (canMoveItemWithinDrive=false)';
    } else if (it.isProtected) {
      it.skipReason = 'PASTA_PROTEGIDA_OU_PROJETO';
    }
  }

  return { folderMap: folderMap, pastasOrganizadasIds: pastasOrganizadasIds };
}

function executarMotorDeduplicacaoPuro_(itens, pastasOrganizadasIds, config) {
  const linhasPlanilha = [];
  const idsJaAgrupadosN1 = new Set();
  let contadorGrupo = 1;
  let bytesEconomiaPotencialN1 = 0;

  const candidatos = itens.filter(it => {
    if (!it.inScope) return false;
    if (it.mimeType === 'application/vnd.google-apps.folder') return false;
    if (it.isShortcut) return false;
    if (it.mimeType.startsWith('application/vnd.google-apps.')) return false;
    if (!it.ownedByMe || it.driveId || it.sharedInherited || !it.canMove) return false;
    if (it.size === null || it.size <= 0) return false;
    if (it.name === config.NOME_PLANILHA_RELATORIO) return false;
    return true;
  });

  const gruposN1 = {};
  for (let i = 0; i < candidatos.length; i++) {
    const f = candidatos[i];
    if (!f.md5Checksum) continue;
    const chave = f.md5Checksum + '__' + f.size;
    if (!gruposN1[chave]) gruposN1[chave] = [];
    gruposN1[chave].push(f);
  }

  Object.keys(gruposN1).forEach(chave => {
    const grupoBruto = gruposN1[chave];
    if (grupoBruto.length < 2) return;

    const subgruposSha = validarSha256QuandoDisponivel_(grupoBruto);
    subgruposSha.forEach(grupo => {
      if (grupo.length < 2) return;

      const idGrupo = 'GRP_' + String(contadorGrupo++).padStart(4, '0');
      const ordenados = ordenarGrupoParaRetencaoPuro_(grupo, pastasOrganizadasIds);

      grupo.forEach(item => idsJaAgrupadosN1.add(item.id));

      const todosProtegidos = grupo.every(item => item.isProtected);

      for (let idx = 0; idx < ordenados.length; idx++) {
        const entrada = ordenados[idx];
        const f = entrada.item;
        const ehMantido = (idx === 0);

        let acaoProposta = 'MANTER_ORIGINAL';
        let economiaBytes = 0;

        if (!ehMantido) {
          if (todosProtegidos || f.isProtected) {
            acaoProposta = 'IGNORAR_PROTEGIDO';
          } else {
            acaoProposta = 'MOVER_PARA_QUARENTENA';
            economiaBytes = f.size || 0;
            bytesEconomiaPotencialN1 += economiaBytes;
          }
        }

        linhasPlanilha.push([
          idGrupo,
          'NIVEL_1_EXATA',
          f.name,
          f.path,
          f.id,
          ehMantido ? 'SIM' : 'NAO',
          entrada.criterio,
          f.md5Checksum,
          f.sha256Checksum || 'N/D',
          f.size,
          ehMantido ? 0 : economiaBytes,
          acaoProposta,
          ''
        ]);
      }
    });
  });

  const restantesN2 = candidatos.filter(f => !idsJaAgrupadosN1.has(f.id));
  const gruposN2 = {};

  for (let i = 0; i < restantesN2.length; i++) {
    const f = restantesN2[i];
    const nomeCan = extrairNomeBaseSemSufixoCopiaPuro_(f.name);
    const chaveN2 = nomeCan + '__' + f.size + '__' + (f.exifTime || 'SEM_EXIF');
    if (!gruposN2[chaveN2]) gruposN2[chaveN2] = [];
    gruposN2[chaveN2].push(f);
  }

  const idsJaAgrupadosN2 = new Set();
  Object.keys(gruposN2).forEach(chave => {
    const grupo = gruposN2[chave];
    if (grupo.length < 2) return;
    if (!grupo.some(f => temPadraoDeCopiaNoNomePuro_(f.name))) return;

    const idGrupo = 'GRP_' + String(contadorGrupo++).padStart(4, '0');
    const ordenados = ordenarGrupoParaRetencaoPuro_(grupo, pastasOrganizadasIds);

    for (let idx = 0; idx < ordenados.length; idx++) {
      const f = ordenados[idx].item;
      idsJaAgrupadosN2.add(f.id);
      linhasPlanilha.push([
        idGrupo,
        'NIVEL_2_PROVAVEL_COPIA',
        f.name,
        f.path,
        f.id,
        idx === 0 ? 'SIM' : 'NAO',
        ordenados[idx].criterio,
        f.md5Checksum || 'N/D',
        f.sha256Checksum || 'N/D',
        f.size,
        0,
        'REVISAR_NAO_MOVER',
        ''
      ]);
    }
  });

  const restantesN3 = candidatos.filter(f =>
    !idsJaAgrupadosN1.has(f.id) &&
    !idsJaAgrupadosN2.has(f.id) &&
    f.mimeType.startsWith('image/') &&
    Boolean(f.exifTime) &&
    Boolean(f.dimensions)
  );

  const gruposN3 = {};
  for (let i = 0; i < restantesN3.length; i++) {
    const f = restantesN3[i];
    const prop = calcularProporcaoImagemPuro_(f.dimensions);
    if (!prop) continue;
    const chaveN3 = f.exifTime + '__' + prop;
    if (!gruposN3[chaveN3]) gruposN3[chaveN3] = [];
    gruposN3[chaveN3].push(f);
  }

  Object.keys(gruposN3).forEach(chave => {
    const grupoBruto = gruposN3[chave];
    if (grupoBruto.length < 2) return;

    const subgruposCamera = separarPorCameraCompativelPuro_(grupoBruto);
    subgruposCamera.forEach(grupo => {
      if (grupo.length < 2) return;

      const idGrupo = 'GRP_' + String(contadorGrupo++).padStart(4, '0');
      const ordenados = ordenarGrupoParaRetencaoPuro_(grupo, pastasOrganizadasIds);

      for (let idx = 0; idx < ordenados.length; idx++) {
        const f = ordenados[idx].item;
        linhasPlanilha.push([
          idGrupo,
          'NIVEL_3_SIMILAR_OU_RAJADA',
          f.name,
          f.path,
          f.id,
          idx === 0 ? 'SIM' : 'NAO',
          ordenados[idx].criterio,
          f.md5Checksum || 'N/D',
          f.sha256Checksum || 'N/D',
          f.size,
          0,
          'SIMILAR_DECISAO_HUMANA',
          ''
        ]);
      }
    });
  });

  return {
    linhasPlanilha: linhasPlanilha,
    bytesEconomiaPotencialN1: bytesEconomiaPotencialN1
  };
}

function ordenarGrupoParaRetencaoPuro_(grupo, pastasOrganizadasIds) {
  const copiaLista = grupo.slice();

  copiaLista.sort((a, b) => {
    if (Boolean(a.isProtected) !== Boolean(b.isProtected)) return a.isProtected ? -1 : 1;

    const aSig = pastasOrganizadasIds.has(a.parentId);
    const bSig = pastasOrganizadasIds.has(b.parentId);
    if (aSig !== bSig) return aSig ? -1 : 1;

    const aLimpo = !temPadraoDeCopiaNoNomePuro_(a.name);
    const bLimpo = !temPadraoDeCopiaNoNomePuro_(b.name);
    if (aLimpo !== bLimpo) return aLimpo ? -1 : 1;

    const tA = Date.parse(a.createdTime || '') || Number.MAX_SAFE_INTEGER;
    const tB = Date.parse(b.createdTime || '') || Number.MAX_SAFE_INTEGER;
    if (tA !== tB) return tA - tB;

    if (Boolean(a.hasStarOrDesc) !== Boolean(b.hasStarOrDesc)) return a.hasStarOrDesc ? -1 : 1;

    return a.id.localeCompare(b.id);
  });

  const vencedor = copiaLista[0];
  return copiaLista.map((item, idx) => {
    if (idx > 0) {
      return { item: item, criterio: 'REPLICA_DE_' + vencedor.id };
    }
    const segundo = copiaLista[1];
    let motivo = '6_MENOR_ID_DETERMINISTICO';
    if (Boolean(vencedor.isProtected) !== Boolean(segundo.isProtected)) {
      motivo = '1_PASTA_PROTEGIDA';
    } else if (pastasOrganizadasIds.has(vencedor.parentId) !== pastasOrganizadasIds.has(segundo.parentId)) {
      motivo = '2_PASTA_NOME_SIGNIFICATIVO';
    } else if ((!temPadraoDeCopiaNoNomePuro_(vencedor.name)) !== (!temPadraoDeCopiaNoNomePuro_(segundo.name))) {
      motivo = '3_NOME_MAIS_LIMPO';
    } else if ((Date.parse(vencedor.createdTime || '') || 0) !== (Date.parse(segundo.createdTime || '') || 0)) {
      motivo = '4_CREATED_TIME_MAIS_ANTIGO';
    } else if (Boolean(vencedor.hasStarOrDesc) !== Boolean(segundo.hasStarOrDesc)) {
      motivo = '5_ESTRELA_OU_DESCRICAO';
    }
    return { item: item, criterio: 'MANTIDO_POR_' + motivo };
  });
}

function temPadraoDeCopiaNoNomePuro_(nome) {
  return /^(c[óo]pia de |copy of )/i.test(nome) || /\(\d+\)(\.[^.]+)?$/i.test(nome) || / - c[óo]pia(\.[^.]+)?$/i.test(nome);
}

function extrairNomeBaseSemSufixoCopiaPuro_(nome) {
  return nome
    .replace(/^(c[óo]pia de |copy of )/i, '')
    .replace(/\s*\(\d+\)(?=\.[^.]+$|$)/i, '')
    .replace(/\s*-\s*c[óo]pia(?=\.[^.]+$|$)/i, '')
    .trim()
    .toLowerCase();
}

function calcularProporcaoImagemPuro_(dimensionsStr) {
  const partes = String(dimensionsStr || '').split('x');
  if (partes.length !== 2) return null;
  const w = Number(partes[0]);
  const h = Number(partes[1]);
  if (!w || !h || h <= 0) return null;
  return (w / h).toFixed(2);
}

function validarSha256QuandoDisponivel_(grupo) {
  const resultado = [];
  for (let i = 0; i < grupo.length; i++) {
    const item = grupo[i];
    let alocado = false;
    for (let j = 0; j < resultado.length; j++) {
      const sub = resultado[j];
      const ref = sub[0];
      if (!item.sha256Checksum || !ref.sha256Checksum || item.sha256Checksum === ref.sha256Checksum) {
        sub.push(item);
        alocado = true;
        break;
      }
    }
    if (!alocado) resultado.push([item]);
  }
  return resultado;
}

function separarPorCameraCompativelPuro_(grupo) {
  const subgrupos = [];
  for (let i = 0; i < grupo.length; i++) {
    const item = grupo[i];
    let colocado = false;
    for (let j = 0; j < subgrupos.length; j++) {
      const sub = subgrupos[j];
      const compativel = sub.every(existente =>
        !item.camera || !existente.camera || item.camera.toLowerCase() === existente.camera.toLowerCase()
      );
      if (compativel) {
        sub.push(item);
        colocado = true;
        break;
      }
    }
    if (!colocado) subgrupos.push([item]);
  }
  return subgrupos;
}

function resolverCaminhoPuro_(parentId, folderMap) {
  if (!parentId || parentId === 'root' || !folderMap[parentId]) return '/Meu Drive';
  const partes = [];
  let atual = parentId;
  const visitados = new Set();
  while (atual && folderMap[atual] && !visitados.has(atual)) {
    visitados.add(atual);
    partes.unshift(folderMap[atual].name);
    atual = folderMap[atual].parentId;
  }
  return '/Meu Drive/' + partes.join('/');
}

function obterAncestralDeTopoDaPasta_(folderId, folderMap, escopoRaizId) {
  let atual = folderId;
  const visitados = new Set();
  while (atual && folderMap[atual] && !visitados.has(atual)) {
    visitados.add(atual);
    const pai = folderMap[atual].parentId;
    if (!pai || pai === 'root' || pai === escopoRaizId || !folderMap[pai]) {
      return atual;
    }
    atual = pai;
  }
  return folderId;
}

function pertenceSubarvore_(itemId, parentId, alvoRaizId, folderMap) {
  if (itemId === alvoRaizId || parentId === alvoRaizId) return true;
  let atual = parentId;
  const visitados = new Set();
  while (atual && folderMap[atual] && !visitados.has(atual)) {
    visitados.add(atual);
    if (atual === alvoRaizId || folderMap[atual].parentId === alvoRaizId) return true;
    atual = folderMap[atual].parentId;
  }
  return false;
}

function testes_logica() {
  Logger.log('Iniciando testes_logica() com dados sintéticos...');

  const itensSinteticos = [
    { id: 'FOLDER_PROJ', name: 'meu-repo', mimeType: 'application/vnd.google-apps.folder', parentId: 'root', ownedByMe: true, shared: false, canMove: true },
    { id: 'FOLDER_ORG', name: 'Documentos_Importantes', mimeType: 'application/vnd.google-apps.folder', parentId: 'root', ownedByMe: true, shared: false, canMove: true },
    { id: 'MARKER_GIT', name: 'package.json', mimeType: 'application/json', parentId: 'FOLDER_PROJ', size: 500, md5Checksum: 'MD5_PKG', ownedByMe: true, shared: false, canMove: true },

    { id: 'DUP_1', name: 'Relatório_Anual (1).pdf', mimeType: 'application/pdf', parentId: 'root', size: 204800, md5Checksum: 'HASH_RELATORIO', sha256Checksum: 'SHA_1', createdTime: '2026-01-10T10:00:00Z', ownedByMe: true, shared: false, canMove: true },
    { id: 'DUP_2', name: 'Relatório_Anual.pdf', mimeType: 'application/pdf', parentId: 'FOLDER_ORG', size: 204800, md5Checksum: 'HASH_RELATORIO', sha256Checksum: 'SHA_1', createdTime: '2026-02-10T10:00:00Z', ownedByMe: true, shared: false, canMove: true },
    { id: 'DUP_3', name: 'Cópia de Relatório_Anual.pdf', mimeType: 'application/pdf', parentId: 'root', size: 204800, md5Checksum: 'HASH_RELATORIO', sha256Checksum: 'SHA_1', createdTime: '2026-01-01T10:00:00Z', ownedByMe: true, shared: false, canMove: true },

    { id: 'BURST_1', name: 'IMG_001.JPG', mimeType: 'image/jpeg', parentId: 'root', size: 4000000, md5Checksum: 'MD5_FOTO_A', exifTime: '2026:10:30 14:20:05', dimensions: '4000x3000', camera: 'Sony A7', createdTime: '2026-10-30T14:20:05Z', ownedByMe: true, shared: false, canMove: true },
    { id: 'BURST_2', name: 'IMG_002.JPG', mimeType: 'image/jpeg', parentId: 'root', size: 4000000, md5Checksum: 'MD5_FOTO_B', exifTime: '2026:10:30 14:20:05', dimensions: '4000x3000', camera: 'Sony A7', createdTime: '2026-10-30T14:20:06Z', ownedByMe: true, shared: false, canMove: true },

    { id: 'NO_EXIF', name: 'WhatsApp_Image.jpeg', mimeType: 'image/jpeg', parentId: 'root', size: 150000, md5Checksum: 'MD5_WA', exifTime: '', dimensions: '1280x720', camera: '', createdTime: '2026-11-01T10:00:00Z', ownedByMe: true, shared: false, canMove: true },

    { id: 'GDOC_1', name: 'Lista de Compras', mimeType: 'application/vnd.google-apps.document', parentId: 'root', size: null, md5Checksum: '', ownedByMe: true, shared: false, canMove: true },
    { id: 'GDOC_2', name: 'Lista de Compras', mimeType: 'application/vnd.google-apps.document', parentId: 'root', size: null, md5Checksum: '', ownedByMe: true, shared: false, canMove: true }
  ];

  const ctx = construirContextoDeArvorePuro_(itensSinteticos, CONFIG);
  const res = executarMotorDeduplicacaoPuro_(itensSinteticos, ctx.pastasOrganizadasIds, CONFIG);

  const itemPkg = itensSinteticos.find(x => x.id === 'MARKER_GIT');
  assert_(itemPkg.isProtected === true, 'Falha: Projeto com package.json deveria estar protegido.');

  const linhasN1 = res.linhasPlanilha.filter(r => r[1] === 'NIVEL_1_EXATA');
  assert_(linhasN1.length === 3, 'Falha: Grupo de 3 duplicatas exatas deve gerar 3 linhas no relatório.');
  const mantidosN1 = linhasN1.filter(r => r[5] === 'SIM');
  assert_(mantidosN1.length === 1 && mantidosN1[0][4] === 'DUP_2', 'Falha: DUP_2 (em pasta significativa) deveria ser o único mantido.');

  const linhasBurst = res.linhasPlanilha.filter(r => r[4] === 'BURST_1' || r[4] === 'BURST_2');
  assert_(linhasBurst.length === 2 && linhasBurst.every(r => r[1] === 'NIVEL_3_SIMILAR_OU_RAJADA' && r[11] === 'SIMILAR_DECISAO_HUMANA'), 'Falha: Fotos em rajada deveriam ser apenas Nível 3 SIMILAR.');

  const linhasGDoc = res.linhasPlanilha.filter(r => r[4] === 'GDOC_1' || r[4] === 'GDOC_2');
  assert_(linhasGDoc.length === 0, 'Falha: Arquivos nativos do Google não devem entrar na deduplicação.');

  Logger.log('SUCESSO: Todos os casos de teste em testes_logica() passaram nas asserções!');
}

function assert_(condicao, mensagemErro) {
  if (!condicao) {
    throw new Error('[ASSERTION FAILED] ' + mensagemErro);
  }
}

function obterOuCriarPlanilhaRelatorio_(props) {
  let spreadsheetId = props.getProperty('RELATORIO_SPREADSHEET_ID');
  let ss = null;

  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    const q = "name = '" + CONFIG.NOME_PLANILHA_RELATORIO + "' and mimeType = 'application/vnd.google-apps.spreadsheet' and 'me' in owners and trashed = false";
    const busca = chamarApiComRetry_(() => Drive.Files.list({ q: q, pageSize: 1, fields: 'files(id, name)' }));
    if (busca.files && busca.files.length > 0) {
      ss = SpreadsheetApp.openById(busca.files[0].id);
    } else {
      ss = SpreadsheetApp.create(CONFIG.NOME_PLANILHA_RELATORIO);
    }
    props.setProperty('RELATORIO_SPREADSHEET_ID', ss.getId());
  }

  inicializarAbasPadrao_(ss);
  return ss;
}

function inicializarAbasPadrao_(ss) {
  const esquemas = [
    {
      nome: 'Inventario',
      colunas: [
        'ID', 'Nome', 'MIME Type', 'Caminho', 'ID da Pasta Pai', 'MD5 Checksum',
        'SHA256 Checksum', 'Tamanho (Bytes)', 'CreatedTime', 'ModifiedTime',
        'EXIF Time', 'Dimensões', 'Câmera', 'Dono (ownedByMe)', 'Compartilhado',
        'DriveID', 'Pode Mover', 'Atalho', 'Estrela/Descrição', 'Protegido', 'Status Elegibilidade'
      ]
    },
    {
      nome: 'Duplicatas_Detectadas',
      colunas: [
        'Grupo', 'Nível (1, 2 ou 3)', 'Arquivo', 'Caminho', 'ID', 'Mantido?',
        'Critério de Retenção', 'MD5', 'SHA256', 'Tamanho (Bytes)',
        'Tamanho Economizado Potencial (Bytes)', 'Ação Proposta', 'Aprovado (SIM ou NÃO)'
      ]
    },
    {
      nome: 'Ignorados',
      colunas: ['ID', 'Nome', 'MIME Type', 'Caminho', 'Motivo de Exclusão do Plano']
    },
    {
      nome: 'Resumo',
      colunas: ['Categoria / Métrica', 'Quantidade / Valor', 'Observações']
    },
    {
      nome: 'Config',
      colunas: ['Chave', 'Valor', 'Descrição']
    },
    {
      nome: 'Log',
      colunas: ['Data e Hora (ISO)', 'Função', 'ID Alvo', 'Ação', 'Resultado / Detalhes']
    }
  ];

  esquemas.forEach((esq, idx) => {
    let aba = ss.getSheetByName(esq.nome);
    if (!aba) {
      aba = (idx === 0 && ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0)
        ? ss.getSheets()[0].setName(esq.nome)
        : ss.insertSheet(esq.nome);
    }
    if (aba.getLastRow() === 0) {
      aba.getRange(1, 1, 1, esq.colunas.length).setValues([esq.colunas]).setFontWeight('bold');
    }
  });

  const abaCfg = ss.getSheetByName('Config');
  if (abaCfg.getLastRow() <= 1) {
    abaCfg.getRange(2, 1, 5, 3).setValues([
      ['DRY_RUN', String(CONFIG.DRY_RUN), 'Modo simulação ativo por padrão'],
      ['PERFIL_ESCOPO', CONFIG.PERFIL_ESCOPO, 'Perfil definido na Etapa 1'],
      ['PASTA_RAIZ_DO_ESCOPO', CONFIG.PASTA_RAIZ_DO_ESCOPO || 'ROOT', 'Escopo da auditoria'],
      ['FRASE_DE_CONFIRMACAO', '', 'Preencha na Etapa 4 somente quando for executar'],
      ['TIMESTAMP_ULTIMA_AUDITORIA', new Date().toISOString(), 'Validade recomendada: 24h']
    ]);
  }
}

function escreverResumo_(abaRes, itens, ignoradosRows, resultadoDedup) {
  limparDadosMantendoCabecalho_(abaRes);

  const contagemPorMotivo = {};
  (ignoradosRows || []).forEach(r => {
    const m = r[4];
    contagemPorMotivo[m] = (contagemPorMotivo[m] || 0) + 1;
  });

  const dupRows = (resultadoDedup && resultadoDedup.linhasPlanilha) ? resultadoDedup.linhasPlanilha : [];
  const n1Replicas = dupRows.filter(r => r[1] === 'NIVEL_1_EXATA' && r[5] === 'NAO').length;
  const n2Total = dupRows.filter(r => r[1] === 'NIVEL_2_PROVAVEL_COPIA').length;
  const n3Total = dupRows.filter(r => r[1] === 'NIVEL_3_SIMILAR_OU_RAJADA').length;
  const mbEconomia = (((resultadoDedup && resultadoDedup.bytesEconomiaPotencialN1) || 0) / (1024 * 1024)).toFixed(2);

  const linhasResumo = [
    ['Total de Itens Inventariados (Dono = Eu, Fora da Lixeira)', Array.isArray(itens) ? itens.length : 0, 'Total bruto lido pela API'],
    ['Total de Itens Ignorados / Protegidos', (ignoradosRows || []).length, 'Listados na aba Ignorados'],
    ['Réplicas Exatas Elegíveis à Quarentena (Nível 1)', n1Replicas, 'Exigem Aprovado = SIM linha a linha'],
    ['Itens Sinalizados como Provável Cópia (Nível 2)', n2Total, 'Apenas revisão humana (não movidos)'],
    ['Fotos Sinalizadas como Similares / Rajada (Nível 3)', n3Total, 'Apenas revisão humana (não movidas)'],
    ['Espaço Potencial Economizável em Nível 1 (MB)', mbEconomia, 'Potencial após revisão da quarentena']
  ];

  Object.keys(contagemPorMotivo).forEach(motivo => {
    linhasResumo.push(['Ignorados — ' + motivo, contagemPorMotivo[motivo], 'Bloqueio preventivo de segurança']);
  });

  abaRes.getRange(2, 1, linhasResumo.length, 3).setValues(linhasResumo);
}

function despejarBufferNaAba_(aba, linhas) {
  if (!linhas || linhas.length === 0) return;
  const ss = aba.getParent();
  const totalCelulasEstimado = ss.getSheets().reduce((acc, s) => acc + (s.getLastRow() * s.getLastColumn()), 0);
  if (totalCelulasEstimado + (linhas.length * linhas[0].length) > CONFIG.LIMITE_CELULAS_PLANILHA) {
    throw new Error('Limite de segurança de células do Google Sheets (9M) atingido. Divida o escopo usando PASTA_RAIZ_DO_ESCOPO.');
  }
  const proxLinha = aba.getLastRow() + 1;
  aba.getRange(proxLinha, 1, linhas.length, linhas[0].length).setValues(linhas);
}

function regravaMatrizEmLotes_(aba, linhaInicial, colInicial, matriz) {
  if (!matriz || matriz.length === 0) return;
  const numCols = matriz[0].length;
  const lote = CONFIG.TAMANHO_LOTE_ESCRITA;
  for (let i = 0; i < matriz.length; i += lote) {
    const fatia = matriz.slice(i, i + lote);
    aba.getRange(linhaInicial + i, colInicial, fatia.length, numCols).setValues(fatia);
  }
}

function limparDadosMantendoCabecalho_(aba) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha > 1) {
    aba.getRange(2, 1, ultimaLinha - 1, aba.getLastColumn()).clearContent();
  }
}

function chamarApiComRetry_(fnChamada) {
  const maxTentativas = 5;
  for (let tentativa = 1; tentativa <= maxTentativas; tentativa++) {
    try {
      return fnChamada();
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      const transiente = /(429|500|503|rateLimitExceeded|userRateLimitExceeded|quotaExceeded|backendError)/i.test(msg);
      if (!transiente || tentativa === maxTentativas) {
        throw e;
      }
      const backoffMs = Math.pow(2, tentativa) * 1000 + Math.floor(Math.random() * 500);
      Utilities.sleep(backoffMs);
    }
  }
}

function agendarGatilhoContinuacao_(nomeFuncao) {
  limparGatilhosContinuacao_();
  ScriptApp.newTrigger(nomeFuncao)
    .timeBased()
    .after(60 * 1000)
    .create();
}

function limparGatilhosContinuacao_() {
  const gatilhos = ScriptApp.getProjectTriggers();
  for (let i = 0; i < gatilhos.length; i++) {
    if (gatilhos[i].getHandlerFunction() === 'etapa1_auditoria') {
      ScriptApp.deleteTrigger(gatilhos[i]);
    }
  }
}

function limparEstadoEGatilhos_(props) {
  limparGatilhosContinuacao_();
  props.deleteProperty('AUDITORIA_FASE');
  props.deleteProperty('CURSOR_PAGE_TOKEN');
  props.deleteProperty('TOTAL_COLETADO');
}

function registrarLog_(ss, funcao, idAlvo, acao, resultado) {
  try {
    const abaLog = ss.getSheetByName('Log');
    if (abaLog) {
      const prox = abaLog.getLastRow() + 1;
      abaLog.getRange(prox, 1, 1, 5).setValues([[
        new Date().toISOString(), funcao, idAlvo, acao, resultado
      ]]);
    }
  } catch (e) {
    Logger.log('Erro ao gravar log na planilha: ' + e.message);
  }
}
