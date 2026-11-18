/**
 * ============================================================================
 * PROJETO A (PARTE 2): ETAPA 3 — REGRAS DE ROTEAMENTO
 * ============================================================================
 * Escopo: Somente Leitura no Drive (drive.readonly).
 * Lê as abas 'Inventario' e 'Duplicatas_Detectadas' e gera:
 * - 'Plano_Organizacao' (com coluna Aprovado vazia)
 * - 'Sugestoes_Renomeacao' (opcional, baseado 100% em metadados, sem ler conteúdo)
 */

function etapa3_gerarPlanoOrganizacao() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Execução abortada: outra instância do script está em andamento.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  let ss = null;

  try {
    ss = obterOuCriarPlanilhaRelatorio_(props);
    registrarLog_(ss, 'etapa3_gerarPlanoOrganizacao', 'SISTEMA', 'INICIO', 'Iniciando geração do Plano de Organização.');

    const abaInv = ss.getSheetByName('Inventario');
    const abaDup = ss.getSheetByName('Duplicatas_Detectadas');
    const abaCfg = ss.getSheetByName('Config');

    if (!abaInv || abaInv.getLastRow() <= 1) {
      throw new Error('A aba Inventario está vazia. Execute etapa1_auditoria() antes da Etapa 3.');
    }

    const perfilAtivo = lerPerfilDaAbaConfig_(abaCfg) || CONFIG.PERFIL_ESCOPO || 'CONSERVADOR';
    const idsReplicasQuarentena = coletarIdsReplicasNivel1_(abaDup);

    const ultimaLinhaInv = abaInv.getLastRow();
    const dadosInv = abaInv.getRange(2, 1, ultimaLinhaInv - 1, 21).getValues();
    const itensInventario = dadosInv.map(row => ({
      id: String(row[0] || ''),
      name: String(row[1] || ''),
      mimeType: String(row[2] || ''),
      path: String(row[3] || '/Meu Drive'),
      parentId: String(row[4] || 'root'),
      md5Checksum: String(row[5] || ''),
      size: row[7] !== '' ? Number(row[7]) : null,
      createdTime: String(row[8] || ''),
      modifiedTime: String(row[9] || ''),
      exifTime: String(row[10] || ''),
      dimensions: String(row[11] || ''),
      ownedByMe: row[13] === 'SIM',
      shared: row[14] === 'SIM',
      driveId: String(row[15] || ''),
      canMove: row[16] === 'SIM',
      isShortcut: row[17] === 'SIM',
      isProtected: row[19] === 'SIM',
      eligibilityStatus: String(row[20] || '')
    }));

    const resultadoRoteamento = gerarPlanoRoteamentoPuro_(
      itensInventario,
      idsReplicasQuarentena,
      perfilAtivo,
      CONFIG
    );

    validarInvariantesDoPlanoPuro_(resultadoRoteamento.linhasPlano, idsReplicasQuarentena, itensInventario);

    const abaPlano = inicializarAbaPlanoOrganizacao_(ss);
    limparDadosMantendoCabecalho_(abaPlano);
    if (resultadoRoteamento.linhasPlano.length > 0) {
      regravaMatrizEmLotes_(abaPlano, 2, 1, resultadoRoteamento.linhasPlano);
    }

    const abaRenomear = inicializarAbaSugestoesRenomeacao_(ss);
    limparDadosMantendoCabecalho_(abaRenomear);
    if (resultadoRoteamento.sugestoesRenomeacao.length > 0) {
      regravaMatrizEmLotes_(abaRenomear, 2, 1, resultadoRoteamento.sugestoesRenomeacao);
    }

    atualizarResumoComEtapa3_(ss, resultadoRoteamento.estatisticas, perfilAtivo);

    registrarLog_(
      ss,
      'etapa3_gerarPlanoOrganizacao',
      'SISTEMA',
      'CONCLUIDO',
      'Plano gerado: ' + resultadoRoteamento.linhasPlano.length + ' itens | Perfil: ' + perfilAtivo
    );
    Logger.log('Etapa 3 concluída! Total de linhas em Plano_Organizacao: ' + resultadoRoteamento.linhasPlano.length);
  } catch (erro) {
    if (ss) {
      registrarLog_(ss, 'etapa3_gerarPlanoOrganizacao', 'ERRO_FATAL', 'FALHA', erro.message);
    }
    throw erro;
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function gerarPlanoRoteamentoPuro_(itens, idsReplicasQuarentena, perfilEscopo, config) {
  const linhasPlano = [];
  const sugestoesRenomeacao = [];
  const prefixosTaxonomia = [
    '01_Pessoal_e_Documentos',
    '02_Trabalho_e_Carreira',
    '03_Estudos_e_Academico',
    '04_Tecnologia_e_Projetos',
    '05_Midia_e_Fotos',
    '99_Arquivo_Morto_e_Backups',
    '_TRIAGEM_MANUAL',
    '_QUARENTENA_DUPLICATAS_PARA_REVISAO',
    '_RESTAURADOS'
  ];
  const pastasTriagemInbox = new Set(['downloads', 'temp', 'novos uploads', 'triagem', 'para_organizar']);
  const nomesProtegidosFixos = new Set((config.PASTAS_PROTEGIDAS_NOMES || []).map(n => n.toLowerCase()));

  const stats = {
    altaConfianca: 0,
    mediaConfianca: 0,
    baixaConfianca: 0,
    fotosComExif: 0,
    fotosSemExif: 0,
    pastasProjetoRaiz: 0
  };

  for (let i = 0; i < itens.length; i++) {
    const it = itens[i];

    if (it.name === config.NOME_PLANILHA_RELATORIO) continue;
    if (idsReplicasQuarentena.has(it.id)) continue;
    if (!it.ownedByMe || it.shared || it.driveId || !it.canMove || it.isShortcut) continue;
    if (jaEstaNaTaxonomiaAlvoPuro_(it.path, it.name, prefixosTaxonomia)) continue;

    const estaNaRaiz = (it.path === '/Meu Drive');
    const nomePastaPai = extrairNomePastaFinalDoCaminho_(it.path).toLowerCase();
    const estaEmInbox = pastasTriagemInbox.has(nomePastaPai);

    if (it.mimeType === 'application/vnd.google-apps.folder') {
      if (estaNaRaiz && it.isProtected && !nomesProtegidosFixos.has(it.name.toLowerCase())) {
        const destinoProj = '/Meu Drive/04_Tecnologia_e_Projetos/Repositorios_e_Codigo';
        if (it.path !== destinoProj) {
          linhasPlano.push([
            it.id,
            it.name,
            it.path,
            destinoProj,
            'PASTA_RAIZ_DE_PROJETO_INTACTA (conteúdo interno blindado)',
            'MEDIA',
            'REVISAR_MOVER_PASTA_PROJETO',
            ''
          ]);
          stats.mediaConfianca++;
          stats.pastasProjetoRaiz++;
        }
      }
      continue;
    }

    if (it.isProtected || it.eligibilityStatus !== 'ELEGIVEL') continue;

    const perfil = String(perfilEscopo || 'CONSERVADOR').toUpperCase();
    if (perfil === 'CONSERVADOR' && !estaNaRaiz) continue;
    if (perfil === 'EQUILIBRADO' && !estaNaRaiz && !estaEmInbox) continue;

    const decisao = classificarArquivoPuro_(it);
    if (!decisao || decisao.pastaSugerida === it.path) continue;

    if (decisao.confianca === 'ALTA') stats.altaConfianca++;
    else if (decisao.confianca === 'MEDIA') stats.mediaConfianca++;
    else stats.baixaConfianca++;

    if (decisao.subtipoEstatistica === 'FOTO_COM_EXIF') stats.fotosComExif++;
    if (decisao.subtipoEstatistica === 'FOTO_SEM_EXIF') stats.fotosSemExif++;

    linhasPlano.push([
      it.id,
      it.name,
      it.path,
      decisao.pastaSugerida,
      decisao.regraAplicada,
      decisao.confianca,
      decisao.acao,
      ''
    ]);

    const sugestaoNome = gerarSugestaoRenomeacaoPorMetadadosPuro_(it, decisao);
    if (sugestaoNome) {
      sugestoesRenomeacao.push(sugestaoNome);
    }
  }

  return {
    linhasPlano: linhasPlano,
    sugestoesRenomeacao: sugestoesRenomeacao,
    estatisticas: stats
  };
}

function classificarArquivoPuro_(item) {
  const nome = String(item.name || '');
  const nomeLower = nome.toLowerCase();
  const mime = String(item.mimeType || '').toLowerCase();
  const ext = extrairExtensaoPuro_(nomeLower);

  const avaliacaoSemantica = avaliarRegrasSemanticasPuro_(nomeLower, mime, ext);

  if (mime.startsWith('image/')) {
    if (avaliacaoSemantica && avaliacaoSemantica.confianca === 'ALTA') {
      return avaliacaoSemantica;
    }
    return classificarFotoPorExifPuro_(item);
  }

  if (mime.startsWith('video/') || mime.startsWith('audio/')) {
    return {
      pastaSugerida: '/Meu Drive/05_Midia_e_Fotos/Videos_e_Audios',
      regraAplicada: 'MIDIA_AUDIOVISUAL (MIME: ' + mime + ' — sem subpastas de data por ausência de EXIF na API)',
      confianca: 'ALTA',
      acao: 'MOVER_ARQUIVO',
      subtipoEstatistica: 'VIDEO_AUDIO'
    };
  }

  if (avaliacaoSemantica) {
    return avaliacaoSemantica;
  }

  return {
    pastaSugerida: '/Meu Drive/_TRIAGEM_MANUAL',
    regraAplicada: 'BAIXA_CONFIANCA_SEM_PADRAO_CLARO (MIME: ' + (mime || 'desconhecido') + ')',
    confianca: 'BAIXA',
    acao: 'MOVER_TRIAGEM_MANUAL',
    subtipoEstatistica: 'TRIAGEM'
  };
}

function classificarFotoPorExifPuro_(item) {
  const exifRaw = String(item.exifTime || '').trim();
  const matchExif = exifRaw.match(/^(\d{4})[:\-](\d{2})[:\-](\d{2})(?:\s+(\d{2}):(\d{2}):(\d{2}))?/);

  if (matchExif) {
    const ano = matchExif[1];
    const mes = matchExif[2];
    const dia = Number(matchExif[3]);
    const hora = matchExif[4] !== undefined ? Number(matchExif[4]) : 12;

    let notaFuso = '';
    const pertoViradaMes = (dia === 1 && hora < 3) || (dia >= 28 && hora >= 21);
    if (pertoViradaMes) {
      notaFuso = ' [ALERTA_MEIA_NOITE_FUSO: EXIF sem fuso (' + exifRaw + ') próximo à virada de mês]';
    }

    return {
      pastaSugerida: '/Meu Drive/05_Midia_e_Fotos/Anos/' + ano + '/' + ano + '-' + mes,
      regraAplicada: 'FOTO_EXIF_TIME (' + exifRaw + ')' + notaFuso,
      confianca: 'ALTA',
      acao: 'MOVER_ARQUIVO',
      subtipoEstatistica: 'FOTO_COM_EXIF'
    };
  }

  let sugestaoUpload = 'indisponível';
  const matchCreated = String(item.createdTime || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (matchCreated) {
    sugestaoUpload = matchCreated[1] + '-' + matchCreated[2];
  }

  return {
    pastaSugerida: '/Meu Drive/05_Midia_e_Fotos/Sem_Data_Confiavel',
    regraAplicada: 'FOTO_SEM_EXIF (Sugestão informativa por data de upload createdTime: ' + sugestaoUpload + ')',
    confianca: 'MEDIA',
    acao: 'REVISAR',
    subtipoEstatistica: 'FOTO_SEM_EXIF'
  };
}

function avaliarRegrasSemanticasPuro_(nomeLower, mime, ext) {
  const ehDocOuPdfOuPlanilha = (
    mime === 'application/pdf' ||
    mime.startsWith('application/vnd.google-apps.document') ||
    mime.startsWith('application/vnd.google-apps.spreadsheet') ||
    mime.startsWith('application/vnd.google-apps.presentation') ||
    mime.includes('wordprocessingml') ||
    mime.includes('spreadsheetml') ||
    mime.includes('presentationml') ||
    ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'odt', 'txt', 'md', 'rtf', 'xml'].includes(ext)
  );

  const catalogoRegras = [
    {
      idRegra: 'DOC_IDENTIDADE_CIVIL',
      destino: '/Meu Drive/01_Pessoal_e_Documentos/Identidade_e_Civil',
      tokensFortes: ['cnh', 'passaporte', 'certidao_nascimento', 'certidao_casamento', 'titulo_eleitor', 'titulo_eleitoral', 'reservista', 'identidade', 'cin_identidade', 'foto3x4'],
      tokensContexto: ['rg', 'cpf', 'certidao', 'eleitoral', 'cartorio'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'DOC_FINANCEIRO_IMPOSTOS',
      destino: '/Meu Drive/01_Pessoal_e_Documentos/Financeiro_e_Impostos',
      tokensFortes: ['irpf', 'imposto_de_renda', 'darf', 'nota_fiscal', 'nfe', 'nfse', 'comprovante_pagamento', 'comprovante_pix', 'boleto', 'iptu', 'ipva', 'banco_do_brasil', 'bradesco'],
      tokensContexto: ['recibo', 'comprovante', 'fatura', 'extrato', 'contrato', 'aluguel', 'imposto', 'declaracao_ir', 'orcamento'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'DOC_CASA_DOMESTICO',
      destino: '/Meu Drive/01_Pessoal_e_Documentos/Casa_e_Domestico',
      tokensFortes: ['lista_de_compras', 'casa_nova', 'enxoval', 'geladeira', 'microondas', 'micro_ondas', 'eletrodomesticos'],
      tokensContexto: ['compras', 'panelas', 'fogao', 'lavadora', 'garantia', 'orcamento_casa', 'moveis'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'DOC_SAUDE',
      destino: '/Meu Drive/01_Pessoal_e_Documentos/Saude',
      tokensFortes: ['hemograma', 'receita_medica', 'laudo_medico', 'exame_laboratorial', 'carteira_vacinacao', 'ressonancia', 'ultrassom', 'saude1', 'saude2'],
      tokensContexto: ['exame', 'laudo', 'receita', 'vacina', 'consulta', 'unimed', 'odonto', 'atestado', 'saude'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'CARREIRA_CURRICULO',
      destino: '/Meu Drive/02_Trabalho_e_Carreira/Curriculos_e_Portfolio',
      tokensFortes: ['curriculo', 'curriculum', 'curriculum_vitae', 'resume', 'portfolio', 'vaga_back_end', 'preparacao_big_tech'],
      tokensContexto: ['cv', 'linkedin', 'carta_apresentacao'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'CARREIRA_CERTIFICADOS',
      destino: '/Meu Drive/02_Trabalho_e_Carreira/Certificados_e_Diplomas',
      tokensFortes: ['certificado', 'certificadocesar', 'certificadotecnico', 'certificado_tecnico', 'diploma', 'declaracao_conclusao', 'historico_escolar', 'historico', 'certificacao', 'dp_900', 'az900'],
      tokensContexto: ['bootcamp', 'workshop', 'carga_horaria', 'declaracao'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'CARREIRA_FINANCEIRO_PJ_CLT',
      destino: '/Meu Drive/02_Trabalho_e_Carreira/Financeiro_Profissional',
      tokensFortes: ['das_mei', 'ccmei', 'holerite', 'contracheque', 'informe_rendimentos'],
      tokensContexto: ['mei', 'pro_labore', 'cnpj'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'ACADEMICO_ARTIGOS_PESQUISA',
      destino: '/Meu Drive/03_Estudos_e_Academico/Artigos_e_Pesquisas',
      tokensFortes: ['tcc', 'dissertacao', 'mestrado', 'circuit_breaker', 'textbugger', 'ieee', 'acm', 'cin_ufpe', 'ufpe', 'elicitacao_de_requisitos', 'open_finance', 'resiliencia'],
      tokensContexto: ['artigo', 'paper', 'pesquisa', 'ddos', 'iot', 'sbc', 'abnt', 'latex', 'pre_projeto', 'qualificacao', 'matricula'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'ACADEMICO_LIVROS_PDFS',
      destino: '/Meu Drive/03_Estudos_e_Academico/Livros_e_PDFs',
      tokensFortes: ['ebook', 'apostila', 'livro', 'data_science_do_zero', 'teoria_da_computacao', 'sipser', 'compendio_completo'],
      tokensContexto: ['capitulo', 'manual_tecnico', 'bibliografia'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'ACADEMICO_CURSOS_DISCIPLINAS',
      destino: '/Meu Drive/03_Estudos_e_Academico/Cursos_e_Disciplinas',
      tokensFortes: ['lista_exercicios', 'plano_ensino', 'ementa', 'page_rank', 'redes_neurais', 'caixa_morfologica'],
      tokensContexto: ['disciplina', 'aula', 'slides', 'seminario', 'prova', 'anotacoes_aula', 'english'],
      exigeDocOuImagem: true
    },
    {
      idRegra: 'TECH_HARDWARE_SETUP',
      destino: '/Meu Drive/04_Tecnologia_e_Projetos/Hardware_e_Setup',
      tokensFortes: ['braco_articulado', 'bracos_articulados', 'smartwatches', 'f80n', 'se2726h', 'ryzen', 'nvme', 'hub_usb'],
      tokensContexto: ['monitor', 'setup', 'ssd', 'teclado', 'mouse', 'macbook', 'hardware', 'benchmark'],
      exigeDocOuImagem: true
    }
  ];

  const nomeNormalizado = removerAcentosPuro_(nomeLower);
  const matchesCategorias = [];

  for (let i = 0; i < catalogoRegras.length; i++) {
    const regra = catalogoRegras[i];
    const fortesEncontrados = encontrarTokensNoTextoPuro_(nomeNormalizado, regra.tokensFortes);
    const contextoEncontrados = encontrarTokensNoTextoPuro_(nomeNormalizado, regra.tokensContexto);
    const totalTokens = fortesEncontrados.length + contextoEncontrados.length;

    if (totalTokens > 0) {
      matchesCategorias.push({
        regra: regra,
        fortes: fortesEncontrados,
        contexto: contextoEncontrados,
        pontuacao: (fortesEncontrados.length * 2) + contextoEncontrados.length
      });
    }
  }

  if (matchesCategorias.length > 1) {
    matchesCategorias.sort((a, b) => b.pontuacao - a.pontuacao);
    const principal = matchesCategorias[0];
    const secundaria = matchesCategorias[1];
    const todosTokens = principal.fortes.concat(principal.contexto).join(', ');
    return {
      pastaSugerida: principal.regra.destino,
      regraAplicada: 'AMBIGUIDADE_ENTRE_CATEGORIAS (' + principal.regra.idRegra + ' vs ' + secundaria.regra.idRegra + ' | Tokens: ' + todosTokens + ')',
      confianca: 'MEDIA',
      acao: 'REVISAR'
    };
  }

  if (matchesCategorias.length === 1) {
    const match = matchesCategorias[0];
    const todosTokens = match.fortes.concat(match.contexto).join(', ');
    const tipoCompativel = !match.regra.exigeDocOuImagem || ehDocOuPdfOuPlanilha || mime.startsWith('image/');

    if (tipoCompativel && (match.fortes.length >= 1 || match.contexto.length >= 2)) {
      return {
        pastaSugerida: match.regra.destino,
        regraAplicada: match.regra.idRegra + ' (Tokens: [' + todosTokens + '] + MIME/Ext compatível)',
        confianca: 'ALTA',
        acao: 'MOVER_ARQUIVO'
      };
    }

    return {
      pastaSugerida: match.regra.destino,
      regraAplicada: match.regra.idRegra + '_PALAVRA_ISOLADA (Token: [' + todosTokens + '] — requer validação humana)',
      confianca: 'MEDIA',
      acao: 'REVISAR'
    };
  }

  if (['py', 'java', 'go', 'js', 'ts', 'sql', 'sh', 'zsh', 'yaml', 'yml', 'tf', 'ipynb', 'dockerfile'].includes(ext) ||
      mime === 'application/vnd.google-apps.colab') {
    return {
      pastaSugerida: '/Meu Drive/04_Tecnologia_e_Projetos/Scripts_e_Automacoes',
      regraAplicada: 'EXTENSAO_CODIGO_OU_NOTEBOOK (.' + (ext || 'colab') + ')',
      confianca: 'ALTA',
      acao: 'MOVER_ARQUIVO'
    };
  }

  if (['csv', 'tsv', 'parquet', 'jsonl', 'sqlite', 'db', 'h5', 'onnx', 'pt', 'pkl'].includes(ext)) {
    return {
      pastaSugerida: '/Meu Drive/04_Tecnologia_e_Projetos/Datasets_e_Modelos',
      regraAplicada: 'EXTENSAO_DATASET_OU_MODELO (.' + ext + ')',
      confianca: 'ALTA',
      acao: 'MOVER_ARQUIVO'
    };
  }

  if (['iso', 'dmg', 'pkg', 'msi', 'exe', 'deb', 'rpm', 'apk'].includes(ext)) {
    return {
      pastaSugerida: '/Meu Drive/04_Tecnologia_e_Projetos/Backups_Config_e_Instaladores',
      regraAplicada: 'EXTENSAO_INSTALADOR_OU_IMAGEM (.' + ext + ')',
      confianca: 'ALTA',
      acao: 'MOVER_ARQUIVO'
    };
  }

  if (['zip', 'rar', '7z', 'tar', 'gz', 'tgz', 'bak'].includes(ext)) {
    return {
      pastaSugerida: '/Meu Drive/99_Arquivo_Morto_e_Backups',
      regraAplicada: 'ARQUIVO_COMPACTADO_GENERICO (.' + ext + ' — revisar antes de arquivar)',
      confianca: 'MEDIA',
      acao: 'REVISAR'
    };
  }

  return null;
}

function validarInvariantesDoPlanoPuro_(linhasPlano, idsReplicasQuarentena, itensInventario) {
  const idsVistosNoPlano = new Set();
  const mapaInventario = {};
  for (let i = 0; i < itensInventario.length; i++) {
    mapaInventario[itensInventario[i].id] = itensInventario[i];
  }

  for (let i = 0; i < linhasPlano.length; i++) {
    const row = linhasPlano[i];
    const id = row[0];
    const pastaAtual = row[2];
    const pastaSugerida = row[3];
    const aprovadoCol = row[7];

    if (idsVistosNoPlano.has(id)) {
      throw new Error('[VIOLACAO DE INVARIANTE] O FileID ' + id + ' aparece mais de uma vez no Plano_Organizacao.');
    }
    idsVistosNoPlano.add(id);

    if (idsReplicasQuarentena.has(id)) {
      throw new Error('[VIOLACAO DE INVARIANTE] O FileID ' + id + ' já está alocado para Quarentena de Duplicatas e não pode receber segunda ação.');
    }

    if (pastaAtual === pastaSugerida) {
      throw new Error('[VIOLACAO DE INVARIANTE] Origem igual ao destino para o FileID ' + id + ' (' + pastaAtual + ').');
    }

    if (aprovadoCol !== '') {
      throw new Error('[VIOLACAO DE INVARIANTE] A coluna Aprovado deve iniciar vazia para revisão humana.');
    }

    const meta = mapaInventario[id];
    if (meta && (!meta.ownedByMe || meta.shared || meta.driveId || meta.isShortcut || !meta.canMove)) {
      throw new Error('[VIOLACAO DE INVARIANTE] Item bloqueado por permissão/compartilhamento entrou no plano: ' + id);
    }
  }
}

function gerarSugestaoRenomeacaoPorMetadadosPuro_(item, decisao) {
  const nome = String(item.name || '');
  const padraoCaotico = /^(scan\d*|scanned_\d+.*|img_\d+|dsc_\d+|whatsapp image.*|documento sem t[íi]tulo|planilha sem t[íi]tulo|apresenta[çc][ãa]o sem t[íi]tulo|crie um docs.*|untitled.*)$/i;

  if (!padraoCaotico.test(nome.replace(/\.[^.]+$/, '').trim())) {
    return null;
  }

  const ext = extrairExtensaoPuro_(nome);
  const sufixoExt = ext ? ('.' + ext) : '';
  const dataRef = (item.exifTime || item.createdTime || '').slice(0, 10).replace(/:/g, '-') || 'SEM-DATA';
  const categoriaCurta = decisao.pastaSugerida.split('/').pop() || 'Item';
  const nomeLimpoSugerido = dataRef + '_' + categoriaCurta + '_' + item.id.slice(0, 6) + sufixoExt;

  return [
    item.id,
    item.name,
    nomeLimpoSugerido,
    item.path,
    'Baseado apenas em Data (' + dataRef + ') + Categoria (' + categoriaCurta + ') — sem leitura de conteúdo',
    ''
  ];
}

function encontrarTokensNoTextoPuro_(textoNormalizado, listaTokens) {
  const encontrados = [];
  for (let i = 0; i < listaTokens.length; i++) {
    const token = removerAcentosPuro_(listaTokens[i].toLowerCase()).replace(/_/g, '[\\s_\\-]');
    const regex = new RegExp('(?:^|[\\s_\\-.,()\\[\\]])(' + token + ')(?:$|[\\s_\\-.,()\\[\\]])', 'i');
    if (regex.test(textoNormalizado)) {
      encontrados.push(listaTokens[i]);
    }
  }
  return encontrados;
}

function removerAcentosPuro_(str) {
  return String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function extrairExtensaoPuro_(nomeArquivo) {
  const partes = String(nomeArquivo || '').split('.');
  if (partes.length < 2) return '';
  return partes.pop().toLowerCase().trim();
}

function extrairNomePastaFinalDoCaminho_(pathStr) {
  const partes = String(pathStr || '').split('/').filter(Boolean);
  return partes.length > 0 ? partes[partes.length - 1] : 'Meu Drive';
}

function jaEstaNaTaxonomiaAlvoPuro_(pathStr, itemName, prefixosTaxonomia) {
  for (let i = 0; i < prefixosTaxonomia.length; i++) {
    const pref = prefixosTaxonomia[i];
    if (pathStr.indexOf('/Meu Drive/' + pref) === 0 || itemName === pref) {
      return true;
    }
  }
  return false;
}

function testes_logica_roteamento() {
  Logger.log('Iniciando testes_logica_roteamento() com casos sintéticos...');

  const idsReplicasQuarentena = new Set(['ID_REPLICA_N1']);

  const itensTeste = [
    {
      id: 'FOTO_EXIF_OK', name: 'IMG_2040.JPG', mimeType: 'image/jpeg', path: '/Meu Drive',
      parentId: 'root', exifTime: '2024:07:15 14:30:00', createdTime: '2025-01-01T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'FOTO_EXIF_MEIA_NOITE', name: 'IMG_VIRADA.JPG', mimeType: 'image/jpeg', path: '/Meu Drive',
      parentId: 'root', exifTime: '2026:10:31 23:45:10', createdTime: '2026-11-01T02:45:10Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'FOTO_SEM_EXIF', name: 'WhatsApp Image 2026-11-05.jpeg', mimeType: 'image/jpeg', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-11-05T18:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'DOC_IRPF', name: 'Comprovante_IRPF_2026.pdf', mimeType: 'application/pdf', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-05-01T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'DOC_ISOLADO', name: 'recibo_simples.pdf', mimeType: 'application/pdf', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-05-01T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'DOC_OPACO', name: 'scan00239.pdf', mimeType: 'application/pdf', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-08-10T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'ID_REPLICA_N1', name: 'Comprovante_IRPF_2026 (1).pdf', mimeType: 'application/pdf', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-05-01T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    },
    {
      id: 'ARGO_YAML', name: 'argocd_config.yaml', mimeType: 'text/yaml', path: '/Meu Drive',
      parentId: 'root', exifTime: '', createdTime: '2026-09-01T10:00:00Z',
      ownedByMe: true, shared: false, driveId: '', canMove: true, isShortcut: false, isProtected: false, eligibilityStatus: 'ELEGIVEL'
    }
  ];

  const resultado = gerarPlanoRoteamentoPuro_(itensTeste, idsReplicasQuarentena, 'CONSERVADOR', CONFIG);
  validarInvariantesDoPlanoPuro_(resultado.linhasPlano, idsReplicasQuarentena, itensTeste);

  const mapaPlano = {};
  resultado.linhasPlano.forEach(r => { mapaPlano[r[0]] = r; });

  assert_(
    mapaPlano['FOTO_EXIF_OK'][3] === '/Meu Drive/05_Midia_e_Fotos/Anos/2024/2024-07' &&
    mapaPlano['FOTO_EXIF_OK'][5] === 'ALTA',
    'Falha no roteamento de foto com EXIF.'
  );

  assert_(
    mapaPlano['FOTO_EXIF_MEIA_NOITE'][4].indexOf('ALERTA_MEIA_NOITE_FUSO') !== -1,
    'Falha ao documentar alerta de meia-noite sem fuso no EXIF.'
  );

  assert_(
    mapaPlano['FOTO_SEM_EXIF'][3] === '/Meu Drive/05_Midia_e_Fotos/Sem_Data_Confiavel' &&
    mapaPlano['FOTO_SEM_EXIF'][5] === 'MEDIA' &&
    mapaPlano['FOTO_SEM_EXIF'][6] === 'REVISAR',
    'Falha no roteamento de foto sem EXIF.'
  );

  assert_(
    mapaPlano['DOC_IRPF'][3] === '/Meu Drive/01_Pessoal_e_Documentos/Financeiro_e_Impostos' &&
    mapaPlano['DOC_IRPF'][5] === 'ALTA',
    'Falha no roteamento de IRPF.'
  );

  assert_(
    mapaPlano['DOC_ISOLADO'][5] === 'MEDIA' && mapaPlano['DOC_ISOLADO'][6] === 'REVISAR',
    'Falha: palavra isolada deveria receber confiança MEDIA e ação REVISAR.'
  );

  assert_(
    mapaPlano['DOC_OPACO'][3] === '/Meu Drive/_TRIAGEM_MANUAL' &&
    mapaPlano['DOC_OPACO'][5] === 'BAIXA',
    'Falha no roteamento de arquivo opaco para _TRIAGEM_MANUAL.'
  );

  assert_(
    !mapaPlano['ID_REPLICA_N1'],
    'Falha: réplica de Nível 1 não poderia entrar no Plano_Organizacao.'
  );

  assert_(
    mapaPlano['ARGO_YAML'][3] === '/Meu Drive/04_Tecnologia_e_Projetos/Scripts_e_Automacoes',
    'Falha: argocd_config.yaml colidiu indevidamente com outra regra.'
  );

  Logger.log('SUCESSO: Todos os 8 testes de roteamento e invariantes passaram em testes_logica_roteamento()!');
}

function coletarIdsReplicasNivel1_(abaDup) {
  const ids = new Set();
  if (!abaDup || abaDup.getLastRow() <= 1) return ids;
  const dados = abaDup.getRange(2, 1, abaDup.getLastRow() - 1, 12).getValues();
  for (let i = 0; i < dados.length; i++) {
    const nivel = String(dados[i][1] || '');
    const id = String(dados[i][4] || '');
    const mantido = String(dados[i][5] || '');
    const acao = String(dados[i][11] || '');
    if (nivel === 'NIVEL_1_EXATA' && mantido === 'NAO' && acao === 'MOVER_PARA_QUARENTENA' && id) {
      ids.add(id);
    }
  }
  return ids;
}

function lerPerfilDaAbaConfig_(abaCfg) {
  if (!abaCfg || abaCfg.getLastRow() <= 1) return null;
  const linhas = abaCfg.getRange(2, 1, abaCfg.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < linhas.length; i++) {
    if (String(linhas[i][0]).trim() === 'PERFIL_ESCOPO') {
      return String(linhas[i][1]).trim().toUpperCase();
    }
  }
  return null;
}

function inicializarAbaPlanoOrganizacao_(ss) {
  let aba = ss.getSheetByName('Plano_Organizacao');
  if (!aba) {
    aba = ss.insertSheet('Plano_Organizacao');
  }
  const cabecalho = [[
    'ID',
    'Arquivo',
    'Pasta Atual',
    'Pasta Sugerida',
    'Regra Aplicada',
    'Confiança (ALTA / MEDIA / BAIXA)',
    'Ação',
    'Aprovado (SIM ou NÃO)'
  ]];
  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho).setFontWeight('bold');
  return aba;
}

function inicializarAbaSugestoesRenomeacao_(ss) {
  let aba = ss.getSheetByName('Sugestoes_Renomeacao');
  if (!aba) {
    aba = ss.insertSheet('Sugestoes_Renomeacao');
  }
  const cabecalho = [[
    'ID',
    'Nome Atual',
    'Nome Sugerido (Via Metadados)',
    'Pasta Atual',
    'Critério de Metadados',
    'Aprovado (SIM ou NÃO)'
  ]];
  aba.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho).setFontWeight('bold');
  return aba;
}

function atualizarResumoComEtapa3_(ss, stats, perfilAtivo) {
  const abaRes = ss.getSheetByName('Resumo');
  if (!abaRes) return;
  const novasLinhas = [
    ['[ETAPA 3] Perfil de Escopo Aplicado', perfilAtivo, 'Configurável na aba Config'],
    ['[ETAPA 3] Roteamentos de Confiança ALTA (MOVER_ARQUIVO)', stats.altaConfianca, 'Elegíveis para movimentação mediante Aprovado = SIM'],
    ['[ETAPA 3] Roteamentos de Confiança MEDIA (REVISAR)', stats.mediaConfianca, 'Inclui fotos sem EXIF, compactados e termos isolados'],
    ['[ETAPA 3] Roteamentos de Confiança BAIXA (_TRIAGEM_MANUAL)', stats.baixaConfianca, 'Arquivos com nomes genéricos/opacos'],
    ['[ETAPA 3] Fotos com Data EXIF Confiável', stats.fotosComExif, 'Roteadas para 05_Midia_e_Fotos/Anos/AAAA/AAAA-MM'],
    ['[ETAPA 3] Fotos sem EXIF (Enviadas para Sem_Data_Confiavel)', stats.fotosSemExif, 'Nunca usam createdTime como data de captura'],
    ['[ETAPA 3] Pastas-Raiz de Projetos Sugeridas Intactas', stats.pastasProjetoRaiz, 'Conteúdo interno jamais é desmontado']
  ];
  const prox = abaRes.getLastRow() + 1;
  abaRes.getRange(prox, 1, novasLinhas.length, 3).setValues(novasLinhas);
}
