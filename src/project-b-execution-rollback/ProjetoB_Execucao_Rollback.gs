/**
 * ============================================================================
 * PROJETO B: EXECUÇÃO SEGURA, RECONCILIAÇÃO E ROLLBACK
 * ============================================================================
 * Garantias Arquiteturais:
 * 1. ZERO EXCLUSÃO: Não existe nenhuma chamada de deleção ou lixeira neste código.
 * 2. WRITE-AHEAD LOG (WAL): Grava PENDENTE + flush() em Log_Rollback antes de mover.
 * 3. EXECUÇÃO EM ONDAS: Piloto obrigatório (20 itens) -> Onda 2 (200) -> Restante.
 * 4. REVALIDAÇÃO PRÉ-MOVIMENTO: Confere pai, MD5, tamanho, modifiedTime e permissões.
 * 5. DUPLICATAS NÍVEL 1: Confere que o exemplar Mantido segue intacto antes de mover réplica.
 * 6. DISJUNTOR (CIRCUIT BREAKER): Aborta com 5 erros seguidos ou taxa de erro acima de 5%.
 * 7. RECONCILIAÇÃO E ROLLBACK IDEMPOTENTE: Verifica soma de bytes/contagem e reverte.
 */

const CONFIG_EXECUCAO = {
  DRY_RUN: true,
  SPREADSHEET_ID_RELATORIO: '',
  PASTA_RAIZ_DO_ESCOPO: '',
  TAMANHO_LOTE_PILOTO: 20,
  TAMANHO_ONDA_2: 200,
  LIMITE_ERROS_CONSECUTIVOS: 5,
  TAXA_MAXIMA_ERRO_PERMITIDA: 0.05,
  MINIMO_ITENS_PARA_TAXA_ERRO: 20,
  VALIDADE_MAXIMA_AUDITORIA_HORAS: 24,
  TEMPO_MAXIMO_MS: 260000,
  FRASE_DE_CONFIRMACAO: 'EU REVISEI A PLANILHA E AUTORIZO A EXECUCAO',
  PASTA_QUARENTENA_NOME: '_QUARENTENA_DUPLICATAS_PARA_REVISAO',
  PASTA_RESTAURADOS_NOME: '_RESTAURADOS',
  FILTRO_LOTE_ROLLBACK: 'TODOS'
};

function etapa2_executar() {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    Logger.log('Execução abortada: já existe outra instância de etapa2_executar em andamento.');
    return;
  }

  let ss = null;
  try {
    ss = abrirPlanilhaRelatorioValidada_();
    inicializarAbasProjetoB_(ss);

    const abaCfg = ss.getSheetByName('Config');
    const mapaConfig = lerMapaConfig_(abaCfg);
    const modoDryRun = Boolean(CONFIG_EXECUCAO.DRY_RUN || String(mapaConfig['DRY_RUN']).toLowerCase() === 'true');

    registrarLogGeral_(
      ss,
      'etapa2_executar',
      'SISTEMA',
      'INICIO',
      'Modo: ' + (modoDryRun ? 'SIMULACAO (DRY_RUN=true)' : 'PRODUCAO (DRY_RUN=false)')
    );

    const contextoDados = carregarDadosParaExecucao_(ss);

    if (modoDryRun) {
      executarModoSimulacao_(ss, contextoDados);
      limparGatilhosProjetoB_('etapa2_executar');
      registrarLogGeral_(ss, 'etapa2_executar', 'SISTEMA', 'SIMULACAO_CONCLUIDA', 'Aba Simulacao preenchida com sucesso. Nenhum arquivo foi tocado.');
      Logger.log('Simulação concluída com sucesso! Revise a aba "Simulacao" na planilha.');
      return;
    }

    validarTravasDeProducaoPreVoo_(ss, mapaConfig);

    const infoOnda = determinarOndaDeExecucaoAtual_(abaCfg, mapaConfig);
    if (!infoOnda.liberado) {
      limparGatilhosProjetoB_('etapa2_executar');
      registrarLogGeral_(ss, 'etapa2_executar', 'PORTAO_DE_ONDA', 'AGUARDANDO_CONFIRMACAO', infoOnda.mensagem);
      Logger.log('PAUSA DE SEGURANÇA (PORTÃO DE ONDA): ' + infoOnda.mensagem);
      return;
    }

    const resultadoOnda = processarFilaEmProducao_(ss, abaCfg, mapaConfig, contextoDados, infoOnda, startTime);

    if (resultadoOnda.pausadoPorTempo) {
      agendarGatilhoProjetoB_('etapa2_executar');
      registrarLogGeral_(ss, 'etapa2_executar', infoOnda.nomeOnda, 'PAUSA_TEMPO', 'Gatilho agendado para continuar a onda atual.');
      return;
    }

    executarReconciliacaoFinal_(ss, contextoDados);
    limparGatilhosProjetoB_('etapa2_executar');
    registrarLogGeral_(ss, 'etapa2_executar', infoOnda.nomeOnda, 'ONDA_FINALIZADA', resultadoOnda.resumo);
    Logger.log(resultadoOnda.resumo);

  } catch (erro) {
    limparGatilhosProjetoB_('etapa2_executar');
    if (ss) {
      registrarLogGeral_(ss, 'etapa2_executar', 'ERRO_CRITICO', 'ABORTADO', erro.message);
    }
    throw erro;
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function etapa3_rollback() {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    Logger.log('Rollback abortado: já existe outra instância em execução.');
    return;
  }

  let ss = null;
  try {
    ss = abrirPlanilhaRelatorioValidada_();
    inicializarAbasProjetoB_(ss);

    const abaCfg = ss.getSheetByName('Config');
    const mapaConfig = lerMapaConfig_(abaCfg);
    const modoDryRun = Boolean(CONFIG_EXECUCAO.DRY_RUN || String(mapaConfig['DRY_RUN']).toLowerCase() === 'true');
    const filtroLote = String(CONFIG_EXECUCAO.FILTRO_LOTE_ROLLBACK || 'TODOS').trim();

    const abaWal = ss.getSheetByName('Log_Rollback');
    if (!abaWal || abaWal.getLastRow() <= 1) {
      Logger.log('Nenhum registro na aba Log_Rollback para reverter.');
      return;
    }

    const dadosWal = abaWal.getRange(2, 1, abaWal.getLastRow() - 1, 10).getValues();
    const linhasSimulacaoRollback = [];
    let revertidosCount = 0;
    let errosConsecutivos = 0;
    let pastaRestauradosId = null;

    for (let i = dadosWal.length - 1; i >= 0; i--) {
      if (Date.now() - startTime > CONFIG_EXECUCAO.TEMPO_MAXIMO_MS) {
        SpreadsheetApp.flush();
        agendarGatilhoProjetoB_('etapa3_rollback');
        registrarLogGeral_(ss, 'etapa3_rollback', 'SISTEMA', 'PAUSA_TEMPO', 'Gatilho agendado para continuar rollback.');
        return;
      }

      const [logId, ondaLote, tipoOp, fileId, nomeItem, paiOrigem, paiDestino, , statusAtual, detalhesAnt] = dadosWal[i];
      if (statusAtual !== 'CONCLUIDO') continue;
      if (filtroLote !== 'TODOS' && ondaLote !== filtroLote) continue;

      if (modoDryRun) {
        const acaoSim = (String(tipoOp) === 'RENOMEAR')
          ? '[' + nomeItem + '] -> [REVERTER_NOME] -> [De "' + paiDestino + '" de volta para "' + paiOrigem + '"]'
          : '[' + nomeItem + '] -> [REVERTER_MOVIMENTO] -> [De ' + paiDestino + ' de volta para ' + paiOrigem + ']';
        linhasSimulacaoRollback.push([
          new Date().toISOString(),
          ondaLote,
          'ROLLBACK_' + tipoOp,
          fileId,
          nomeItem,
          acaoSim,
          'SIMULACAO_ROLLBACK_OK'
        ]);
        continue;
      }

      try {
        const metaAtual = buscarMetadadosAoVivoArquivo_(fileId);
        if (!metaAtual || metaAtual.trashed) {
          abaWal.getRange(i + 2, 9, 1, 2).setValues([['PULADO_ROLLBACK', 'Arquivo não encontrado ou na lixeira']]);
          continue;
        }

        if (String(tipoOp) === 'RENOMEAR') {
          if (metaAtual.name === String(paiOrigem)) {
            abaWal.getRange(i + 2, 9, 1, 2).setValues([['REVERTIDO_IDEMPOTENTE', 'Arquivo já está com o nome original']]);
            continue;
          }
          if (metaAtual.name !== String(paiDestino)) {
            abaWal.getRange(i + 2, 9, 1, 2).setValues([['PULADO_NOME_MUDOU', 'Nome atual difere do registrado em Log_Rollback']]);
            continue;
          }
          chamarApiComRetry_(() => Drive.Files.update({ name: String(paiOrigem) }, String(fileId), { supportsAllDrives: true, fields: 'id,name' }));
          abaWal.getRange(i + 2, 9, 1, 2).setValues([['REVERTIDO', 'Nome restaurado para [' + paiOrigem + '] | Ant: ' + detalhesAnt]]);
          revertidosCount++;
          errosConsecutivos = 0;
          continue;
        }

        const paisAtuais = metaAtual.parents || [];

        if (paisAtuais.indexOf(paiOrigem) !== -1 && paisAtuais.indexOf(paiDestino) === -1) {
          abaWal.getRange(i + 2, 9, 1, 2).setValues([['REVERTIDO_IDEMPOTENTE', 'Item já estava na pasta de origem']]);
          continue;
        }

        if (paisAtuais.indexOf(paiDestino) === -1) {
          abaWal.getRange(i + 2, 9, 1, 2).setValues([['PULADO_MUDOU_DE_LUGAR', 'Item não está mais no destino registrado (' + paiDestino + ')']]);
          continue;
        }

        let destinoRestauracaoId = paiOrigem;
        let statusFinal = 'REVERTIDO';
        let obsFinal = 'Restaurado ao pai original (' + paiOrigem + ')';

        if (!pastaExisteEAcessivel_(paiOrigem)) {
          if (!pastaRestauradosId) {
            pastaRestauradosId = obterOuCriarPastaPrivadaPorCaminho_(
              '/Meu Drive/' + CONFIG_EXECUCAO.PASTA_RESTAURADOS_NOME,
              abaCfg,
              mapaConfig
            );
          }
          destinoRestauracaoId = pastaRestauradosId;
          statusFinal = 'RESTAURADO_EM_PASTA_EMERGENCIA';
          obsFinal = 'AVISO: Pai original (' + paiOrigem + ') não existe mais; movido para /Meu Drive/' + CONFIG_EXECUCAO.PASTA_RESTAURADOS_NOME;
        }

        moverArquivoApiV3ComRetry_(fileId, paiDestino, destinoRestauracaoId);
        abaWal.getRange(i + 2, 9, 1, 2).setValues([[statusFinal, obsFinal + ' | Ant: ' + detalhesAnt]]);
        revertidosCount++;
        errosConsecutivos = 0;

        if (revertidosCount % 15 === 0) SpreadsheetApp.flush();

      } catch (erroItem) {
        errosConsecutivos++;
        abaWal.getRange(i + 2, 9, 1, 2).setValues([['FALHA_ROLLBACK', erroItem.message]]);
        if (errosConsecutivos >= CONFIG_EXECUCAO.LIMITE_ERROS_CONSECUTIVOS) {
          throw new Error('Disjuntor de Rollback acionado após ' + errosConsecutivos + ' erros consecutivos. Último erro: ' + erroItem.message);
        }
      }
    }

    if (modoDryRun && linhasSimulacaoRollback.length > 0) {
      const abaSim = ss.getSheetByName('Simulacao');
      limparDadosMantendoCabecalho_(abaSim);
      regravaMatrizEmLotes_(abaSim, 2, 1, linhasSimulacaoRollback);
      Logger.log('Simulação de Rollback concluída (' + linhasSimulacaoRollback.length + ' itens). Veja a aba Simulacao.');
    } else {
      limparGatilhosProjetoB_('etapa3_rollback');
      registrarLogGeral_(ss, 'etapa3_rollback', 'SISTEMA', 'CONCLUIDO', 'Itens revertidos: ' + revertidosCount);
      Logger.log('Rollback concluído. Total de itens restaurados: ' + revertidosCount);
    }

  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function carregarDadosParaExecucao_(ss) {
  const abaInv = ss.getSheetByName('Inventario');
  const abaDup = ss.getSheetByName('Duplicatas_Detectadas');
  const abaPlano = ss.getSheetByName('Plano_Organizacao');
  const abaWal = ss.getSheetByName('Log_Rollback');

  if (!abaInv || abaInv.getLastRow() <= 1) {
    throw new Error('Aba Inventario vazia ou ausente.');
  }

  const dadosInv = abaInv.getRange(2, 1, abaInv.getLastRow() - 1, 21).getValues();
  const mapaInventario = {};
  let somaBytesInventario = 0;
  let contagemArquivosInventario = 0;

  for (let i = 0; i < dadosInv.length; i++) {
    const r = dadosInv[i];
    const id = String(r[0] || '');
    const mime = String(r[2] || '');
    const sizeNum = r[7] !== '' ? Number(r[7]) : 0;

    if (mime !== 'application/vnd.google-apps.folder') {
      contagemArquivosInventario++;
      somaBytesInventario += sizeNum;
    }

    mapaInventario[id] = {
      id: id,
      name: String(r[1] || ''),
      mimeType: mime,
      path: String(r[3] || ''),
      parentId: String(r[4] || 'root'),
      md5Checksum: String(r[5] || ''),
      sha256Checksum: String(r[6] || ''),
      size: r[7] !== '' ? Number(r[7]) : null,
      createdTime: String(r[8] || ''),
      modifiedTime: String(r[9] || ''),
      ownedByMe: r[13] === 'SIM',
      shared: r[14] === 'SIM',
      driveId: String(r[15] || ''),
      canMove: r[16] === 'SIM',
      isShortcut: r[17] === 'SIM',
      isProtected: r[19] === 'SIM'
    };
  }

  const idsJaProcessadosWal = new Set();
  if (abaWal && abaWal.getLastRow() > 1) {
    const dadosWal = abaWal.getRange(2, 1, abaWal.getLastRow() - 1, 10).getValues();
    for (let i = 0; i < dadosWal.length; i++) {
      const status = String(dadosWal[i][8] || '');
      const fId = String(dadosWal[i][3] || '');
      if (status === 'CONCLUIDO' && fId) {
        idsJaProcessadosWal.add(fId);
      }
    }
  }

  const mapaMantidoPorGrupo = {};
  const filaOperacoes = [];
  const idsNaFila = new Set();

  if (abaDup && abaDup.getLastRow() > 1) {
    const dadosDup = abaDup.getRange(2, 1, abaDup.getLastRow() - 1, 13).getValues();

    for (let i = 0; i < dadosDup.length; i++) {
      const [grupo, nivel, nome, , id, mantido, , md5, , tamanho] = dadosDup[i];
      if (String(nivel) === 'NIVEL_1_EXATA' && String(mantido).toUpperCase() === 'SIM') {
        mapaMantidoPorGrupo[String(grupo)] = {
          id: String(id),
          name: String(nome),
          md5Checksum: String(md5),
          size: Number(tamanho || 0)
        };
      }
    }

    for (let i = 0; i < dadosDup.length; i++) {
      const [grupo, nivel, nome, caminho, id, mantido, , md5, , tamanho, , acaoProposta, aprovado] = dadosDup[i];
      if (String(aprovado).trim().toUpperCase() !== 'SIM') continue;

      if (String(nivel) !== 'NIVEL_1_EXATA' || String(mantido).toUpperCase() === 'SIM' || String(acaoProposta) !== 'MOVER_PARA_QUARENTENA') {
        continue;
      }

      const fileId = String(id).trim();
      if (!fileId || idsJaProcessadosWal.has(fileId) || idsNaFila.has(fileId)) continue;

      filaOperacoes.push({
        tipoOperacao: 'QUARENTENA_NIVEL_1',
        fileId: fileId,
        nome: String(nome),
        caminhoOrigem: String(caminho),
        caminhoDestino: '/Meu Drive/' + CONFIG_EXECUCAO.PASTA_QUARENTENA_NOME,
        grupoDup: String(grupo),
        keeperEsperado: mapaMantidoPorGrupo[String(grupo)] || null,
        md5Esperado: String(md5),
        sizeEsperado: Number(tamanho || 0)
      });
      idsNaFila.add(fileId);
    }
  }

  if (abaPlano && abaPlano.getLastRow() > 1) {
    const dadosPlano = abaPlano.getRange(2, 1, abaPlano.getLastRow() - 1, 8).getValues();
    for (let i = 0; i < dadosPlano.length; i++) {
      const [id, nome, pastaAtual, pastaSugerida, regra, confianca, acao, aprovado] = dadosPlano[i];
      if (String(aprovado).trim().toUpperCase() !== 'SIM') continue;

      const fileId = String(id).trim();
      if (!fileId || idsJaProcessadosWal.has(fileId) || idsNaFila.has(fileId)) continue;
      if (String(pastaAtual) === String(pastaSugerida)) continue;

      filaOperacoes.push({
        tipoOperacao: 'ORGANIZAR_' + String(acao || 'MOVER'),
        fileId: fileId,
        nome: String(nome),
        caminhoOrigem: String(pastaAtual),
        caminhoDestino: String(pastaSugerida),
        regra: String(regra),
        confianca: String(confianca)
      });
      idsNaFila.add(fileId);
    }
  }

  return {
    mapaInventario: mapaInventario,
    filaOperacoes: filaOperacoes,
    idsJaProcessadosWal: idsJaProcessadosWal,
    contagemArquivosInventario: contagemArquivosInventario,
    somaBytesInventario: somaBytesInventario
  };
}

function executarModoSimulacao_(ss, contextoDados) {
  const abaSim = ss.getSheetByName('Simulacao');
  limparDadosMantendoCabecalho_(abaSim);

  const fila = contextoDados.filaOperacoes;
  const linhasSim = [];

  for (let i = 0; i < fila.length; i++) {
    const op = fila[i];
    const snap = contextoDados.mapaInventario[op.fileId];
    const ondaPrevista = (i < CONFIG_EXECUCAO.TAMANHO_LOTE_PILOTO)
      ? 'PILOTO_20'
      : (i < CONFIG_EXECUCAO.TAMANHO_LOTE_PILOTO + CONFIG_EXECUCAO.TAMANHO_ONDA_2 ? 'ONDA_200' : 'ONDA_FINAL');

    let statusValidacao = 'SIMULACAO_VALIDA';
    let detalheMotivo = op.caminhoDestino;

    if (!snap) {
      statusValidacao = 'BLOQUEADO_NAO_CONSTA_INVENTARIO';
      detalheMotivo = 'ID não encontrado na aba Inventario';
    } else if (!snap.ownedByMe || snap.shared || snap.driveId || !snap.canMove || snap.isShortcut) {
      statusValidacao = 'BLOQUEADO_PERMISSAO_OU_COMPARTILHADO';
      detalheMotivo = 'Item compartilhado, de terceiros ou atalho';
    } else if (op.tipoOperacao === 'QUARENTENA_NIVEL_1' && !op.keeperEsperado) {
      statusValidacao = 'BLOQUEADO_SEM_MANTIDO_NO_GRUPO';
      detalheMotivo = 'Grupo ' + op.grupoDup + ' não possui exemplar Mantido válido';
    }

    const descricaoFormatada = '[' + op.nome + '] -> [' + op.tipoOperacao + '] -> [' + detalheMotivo + ']';

    linhasSim.push([
      new Date().toISOString(),
      ondaPrevista,
      op.tipoOperacao,
      op.fileId,
      op.nome,
      descricaoFormatada,
      statusValidacao
    ]);
  }

  if (linhasSim.length > 0) {
    regravaMatrizEmLotes_(abaSim, 2, 1, linhasSim);
  }
}

function determinarOndaDeExecucaoAtual_(abaCfg, mapaConfig) {
  const estadoOnda = String(mapaConfig['ESTADO_ONDA_ATUAL'] || 'PILOTO_20').trim();
  const pilotoConferido = String(mapaConfig['PILOTO_CONFERIDO'] || '').trim().toUpperCase() === 'SIM';
  const onda2Conferida = String(mapaConfig['ONDA_2_CONFERIDA'] || '').trim().toUpperCase() === 'SIM';

  if (estadoOnda === 'PILOTO_20') {
    return {
      liberado: true,
      nomeOnda: 'PILOTO_20',
      limiteItensNestaOnda: CONFIG_EXECUCAO.TAMANHO_LOTE_PILOTO,
      proximoEstadoAoConcluir: 'AGUARDANDO_APROVACAO_PILOTO'
    };
  }

  if (estadoOnda === 'AGUARDANDO_APROVACAO_PILOTO') {
    if (!pilotoConferido) {
      return {
        liberado: false,
        mensagem: 'O Lote Piloto (até 20 itens) foi concluído. Confira os arquivos no Google Drive e, se estiver tudo certo, preencha PILOTO_CONFERIDO = SIM na aba Config para liberar a Onda 2 (200 itens).'
      };
    }
    salvarChaveNaAbaConfig_(abaCfg, mapaConfig, 'ESTADO_ONDA_ATUAL', 'ONDA_200', 'Onda 2 de até 200 itens em andamento');
    return {
      liberado: true,
      nomeOnda: 'ONDA_200',
      limiteItensNestaOnda: CONFIG_EXECUCAO.TAMANHO_ONDA_2,
      proximoEstadoAoConcluir: 'AGUARDANDO_APROVACAO_ONDA_2'
    };
  }

  if (estadoOnda === 'ONDA_200') {
    if (!pilotoConferido) {
      throw new Error('Violação de protocolo: Tentativa de rodar ONDA_200 sem PILOTO_CONFERIDO = SIM.');
    }
    return {
      liberado: true,
      nomeOnda: 'ONDA_200',
      limiteItensNestaOnda: CONFIG_EXECUCAO.TAMANHO_ONDA_2,
      proximoEstadoAoConcluir: 'AGUARDANDO_APROVACAO_ONDA_2'
    };
  }

  if (estadoOnda === 'AGUARDANDO_APROVACAO_ONDA_2') {
    if (!onda2Conferida) {
      return {
        liberado: false,
        mensagem: 'A Onda 2 (até 200 itens) foi concluída. Confira no Google Drive e preencha ONDA_2_CONFERIDA = SIM na aba Config para liberar o lote final restante.'
      };
    }
    salvarChaveNaAbaConfig_(abaCfg, mapaConfig, 'ESTADO_ONDA_ATUAL', 'ONDA_FINAL', 'Onda final liberada');
    return {
      liberado: true,
      nomeOnda: 'ONDA_FINAL',
      limiteItensNestaOnda: Number.MAX_SAFE_INTEGER,
      proximoEstadoAoConcluir: 'TODAS_AS_ONDAS_CONCLUIDAS'
    };
  }

  if (estadoOnda === 'ONDA_FINAL' || estadoOnda === 'TODAS_AS_ONDAS_CONCLUIDAS') {
    if (!pilotoConferido || !onda2Conferida) {
      throw new Error('Violação de protocolo: Tentativa de rodar ONDA_FINAL sem aprovar Piloto e Onda 2.');
    }
    return {
      liberado: true,
      nomeOnda: 'ONDA_FINAL',
      limiteItensNestaOnda: Number.MAX_SAFE_INTEGER,
      proximoEstadoAoConcluir: 'TODAS_AS_ONDAS_CONCLUIDAS'
    };
  }

  throw new Error('ESTADO_ONDA_ATUAL desconhecido na aba Config: ' + estadoOnda);
}

function processarFilaEmProducao_(ss, abaCfg, mapaConfig, contextoDados, infoOnda, startTime) {
  const abaWal = ss.getSheetByName('Log_Rollback');
  const fila = contextoDados.filaOperacoes;

  let processadosNestaOnda = contarItensJaProcessadosNaOnda_(abaWal, infoOnda.nomeOnda);
  let tentativasSessao = 0;
  let errosSessao = 0;
  let errosConsecutivos = 0;

  for (let i = 0; i < fila.length; i++) {
    if (processadosNestaOnda >= infoOnda.limiteItensNestaOnda) {
      salvarChaveNaAbaConfig_(abaCfg, mapaConfig, 'ESTADO_ONDA_ATUAL', infoOnda.proximoEstadoAoConcluir, 'Aguardando conferência humana da onda');
      return {
        pausadoPorTempo: false,
        resumo: 'Onda ' + infoOnda.nomeOnda + ' atingiu o limite de ' + infoOnda.limiteItensNestaOnda + ' itens. Estado atualizado para ' + infoOnda.proximoEstadoAoConcluir + '.'
      };
    }

    if (Date.now() - startTime > CONFIG_EXECUCAO.TEMPO_MAXIMO_MS) {
      SpreadsheetApp.flush();
      return {
        pausadoPorTempo: true,
        resumo: 'Pausa preventiva por tempo durante a onda ' + infoOnda.nomeOnda
      };
    }

    const op = fila[i];
    const snapInv = contextoDados.mapaInventario[op.fileId];
    tentativasSessao++;

    const revalidacao = revalidarItemAntesDeMover_(op, snapInv, abaCfg, mapaConfig);
    if (!revalidacao.valido) {
      registrarLinhaWalCompleta_(
        abaWal,
        infoOnda.nomeOnda,
        op.tipoOperacao,
        op.fileId,
        op.nome,
        snapInv ? snapInv.parentId : 'DESCONHECIDO',
        op.caminhoDestino,
        'PULADO_MUDOU',
        revalidacao.motivo
      );
      continue;
    }

    const logId = 'WAL_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    const linhaWalIndex = gravarWalPendenteAntesDaAcao_(
      abaWal,
      logId,
      infoOnda.nomeOnda,
      op.tipoOperacao,
      op.fileId,
      op.nome,
      revalidacao.parentIdOrigemConfirmado,
      revalidacao.folderIdDestinoConfirmado
    );

    try {
      moverArquivoApiV3ComRetry_(
        op.fileId,
        revalidacao.parentIdOrigemConfirmado,
        revalidacao.folderIdDestinoConfirmado
      );

      atualizarStatusWal_(abaWal, linhaWalIndex, 'CONCLUIDO', 'Movido para ' + op.caminhoDestino);
      processadosNestaOnda++;
      errosConsecutivos = 0;

      if (processadosNestaOnda % 10 === 0) {
        SpreadsheetApp.flush();
      }
    } catch (erroApi) {
      errosSessao++;
      errosConsecutivos++;
      atualizarStatusWal_(abaWal, linhaWalIndex, 'FALHA', 'Erro API: ' + erroApi.message);
      SpreadsheetApp.flush();

      verificarDisjuntorDeErros_(errosConsecutivos, errosSessao, tentativasSessao, erroApi.message);
    }
  }

  salvarChaveNaAbaConfig_(abaCfg, mapaConfig, 'ESTADO_ONDA_ATUAL', infoOnda.proximoEstadoAoConcluir, 'Fila da onda concluída');
  return {
    pausadoPorTempo: false,
    resumo: 'Processamento da onda ' + infoOnda.nomeOnda + ' concluído. Itens movidos nesta onda: ' + processadosNestaOnda
  };
}

function revalidarItemAntesDeMover_(op, snapInv, abaCfg, mapaConfig) {
  if (!snapInv) {
    return { valido: false, motivo: 'Item ausente na aba Inventario' };
  }

  const atual = buscarMetadadosAoVivoArquivo_(op.fileId);
  if (!atual) {
    return { valido: false, motivo: 'Arquivo não encontrado no Drive (inexistente)' };
  }

  const checkPuro = validarIntegridadePreMovimentoPuro_(snapInv, atual);
  if (!checkPuro.valido) {
    return checkPuro;
  }

  if (op.tipoOperacao === 'QUARENTENA_NIVEL_1') {
    if (!op.keeperEsperado || !op.keeperEsperado.id) {
      return { valido: false, motivo: 'PULADO_MANTIDO_INVALIDO: Grupo sem exemplar Mantido registrado' };
    }
    const keeperAoVivo = buscarMetadadosAoVivoArquivo_(op.keeperEsperado.id);
    const checkKeeper = validarIntegridadeDoMantidoPuro_(op.keeperEsperado, keeperAoVivo);
    if (!checkKeeper.valido) {
      return checkKeeper;
    }
  }

  const folderIdDestino = obterOuCriarPastaPrivadaPorCaminho_(op.caminhoDestino, abaCfg, mapaConfig);
  if (folderIdDestino === snapInv.parentId) {
    return { valido: false, motivo: 'Origem igual ao destino' };
  }

  return {
    valido: true,
    parentIdOrigemConfirmado: snapInv.parentId,
    folderIdDestinoConfirmado: folderIdDestino
  };
}

function validarIntegridadePreMovimentoPuro_(snapInv, metaAoVivo) {
  if (!metaAoVivo || metaAoVivo.trashed) {
    return { valido: false, motivo: 'Arquivo está na lixeira ou indisponível' };
  }
  if (!metaAoVivo.ownedByMe) {
    return { valido: false, motivo: 'Arquivo deixou de pertencer ao usuário (ownedByMe=false)' };
  }
  if (metaAoVivo.shared) {
    return { valido: false, motivo: 'Arquivo passou a estar compartilhado (shared=true)' };
  }
  if (metaAoVivo.driveId) {
    return { valido: false, motivo: 'Arquivo está em Drive Compartilhado' };
  }
  if (!metaAoVivo.capabilities || metaAoVivo.capabilities.canMoveItemWithinDrive === false) {
    return { valido: false, motivo: 'Sem permissão canMoveItemWithinDrive' };
  }

  const paiAoVivo = (metaAoVivo.parents && metaAoVivo.parents.length > 0) ? metaAoVivo.parents[0] : 'root';
  if (paiAoVivo !== snapInv.parentId) {
    return { valido: false, motivo: 'Pai mudou desde a auditoria (Esperado: ' + snapInv.parentId + ' | Atual: ' + paiAoVivo + ')' };
  }

  if (snapInv.md5Checksum && String(metaAoVivo.md5Checksum || '') !== snapInv.md5Checksum) {
    return { valido: false, motivo: 'MD5 mudou desde a auditoria' };
  }

  if (snapInv.size !== null && Number(metaAoVivo.size || 0) !== Number(snapInv.size)) {
    return { valido: false, motivo: 'Tamanho (size) mudou desde a auditoria' };
  }

  if (snapInv.modifiedTime && String(metaAoVivo.modifiedTime || '') !== snapInv.modifiedTime) {
    return { valido: false, motivo: 'Data de modificação (modifiedTime) mudou desde a auditoria' };
  }

  return { valido: true };
}

function validarIntegridadeDoMantidoPuro_(keeperEsperado, keeperAoVivo) {
  if (!keeperAoVivo || keeperAoVivo.trashed) {
    return { valido: false, motivo: 'PULADO_MANTIDO_INVALIDO: Exemplar mantido (' + keeperEsperado.id + ') não existe ou está na lixeira' };
  }
  if (String(keeperAoVivo.md5Checksum || '') !== String(keeperEsperado.md5Checksum)) {
    return { valido: false, motivo: 'PULADO_MANTIDO_INVALIDO: MD5 do exemplar mantido mudou' };
  }
  if (Number(keeperAoVivo.size || 0) !== Number(keeperEsperado.size)) {
    return { valido: false, motivo: 'PULADO_MANTIDO_INVALIDO: Tamanho do exemplar mantido mudou' };
  }
  return { valido: true };
}

function verificarDisjuntorDeErros_(errosConsecutivos, totalErros, totalTentativas, ultimoErroMsg) {
  if (errosConsecutivos >= CONFIG_EXECUCAO.LIMITE_ERROS_CONSECUTIVOS) {
    throw new Error(
      '[DISJUNTOR ACIONADO] Execução abortada após ' + errosConsecutivos +
      ' erros consecutivos. Último erro: ' + ultimoErroMsg
    );
  }
  if (totalTentativas >= CONFIG_EXECUCAO.MINIMO_ITENS_PARA_TAXA_ERRO) {
    const taxaAtual = totalErros / totalTentativas;
    if (taxaAtual > CONFIG_EXECUCAO.TAXA_MAXIMA_ERRO_PERMITIDA) {
      throw new Error(
        '[DISJUNTOR ACIONADO] Taxa de erro (' + (taxaAtual * 100).toFixed(2) +
        '%) ultrapassou o limite de 5%. Último erro: ' + ultimoErroMsg
      );
    }
  }
}

function executarReconciliacaoFinal_(ss, contextoDados) {
  const abaRec = ss.getSheetByName('Reconciliacao');
  const abaWal = ss.getSheetByName('Log_Rollback');
  limparDadosMantendoCabecalho_(abaRec);

  const idsInventario = Object.keys(contextoDados.mapaInventario);
  let contagemArquivosDepois = 0;
  let somaBytesDepois = 0;
  let arquivosDesaparecidos = 0;

  const mapaAoVivo = listarSnapshotAtualParaReconciliacao_();

  for (let i = 0; i < idsInventario.length; i++) {
    const id = idsInventario[i];
    const itemInv = contextoDados.mapaInventario[id];
    const itemVivo = mapaAoVivo[id];

    if (!itemVivo) {
      arquivosDesaparecidos++;
      continue;
    }

    if (itemInv.mimeType !== 'application/vnd.google-apps.folder') {
      contagemArquivosDepois++;
      somaBytesDepois += itemVivo.size;
    }
  }

  let movidosVerificadosOk = 0;
  let movidosForaDoDestino = 0;

  if (abaWal && abaWal.getLastRow() > 1) {
    const dadosWal = abaWal.getRange(2, 1, abaWal.getLastRow() - 1, 10).getValues();
    for (let i = 0; i < dadosWal.length; i++) {
      const tipoOp = String(dadosWal[i][2] || '');
      const status = String(dadosWal[i][8] || '');
      const fId = String(dadosWal[i][3] || '');
      const destEsperado = String(dadosWal[i][6] || '');

      if (status === 'CONCLUIDO' && fId && tipoOp !== 'RENOMEAR') {
        const vivo = mapaAoVivo[fId];
        if (vivo && vivo.parents.indexOf(destEsperado) !== -1) {
          movidosVerificadosOk++;
        } else {
          movidosForaDoDestino++;
        }
      }
    }
  }

  const contagemIgual = (contagemArquivosDepois === contextoDados.contagemArquivosInventario);
  const bytesIguais = (somaBytesDepois === contextoDados.somaBytesInventario);
  const reconciliacaoPerfeita = contagemIgual && bytesIguais && arquivosDesaparecidos === 0 && movidosForaDoDestino === 0;

  const linhasReconciliacao = [
    ['Timestamp da Reconciliação', new Date().toISOString(), '-', 'OK'],
    ['Contagem de Arquivos (Não-Pastas)', contextoDados.contagemArquivosInventario, contagemArquivosDepois, contagemIgual ? 'IGUAL (OK)' : 'DIVERGENTE'],
    ['Soma Total de Bytes', contextoDados.somaBytesInventario, somaBytesDepois, bytesIguais ? 'IGUAL (OK)' : 'DIVERGENTE'],
    ['IDs do Inventário Desaparecidos', 0, arquivosDesaparecidos, arquivosDesaparecidos === 0 ? 'ZERO PERDAS (OK)' : 'ALERTA_CRITICO'],
    ['Itens Movidos Confirmados no Destino', movidosVerificadosOk, movidosVerificadosOk, movidosForaDoDestino === 0 ? '100% NO DESTINO (OK)' : ('DIVERGENTES: ' + movidosForaDoDestino)],
    ['VEREDITO GERAL DA RECONCILIAÇÃO', 'INTEGRIDADE 100%', reconciliacaoPerfeita ? 'APROVADO' : 'REVISAR_DIVERGENCIAS', reconciliacaoPerfeita ? 'OK' : 'FALHA_VERIFICACAO']
  ];

  abaRec.getRange(2, 1, linhasReconciliacao.length, 4).setValues(linhasReconciliacao);
}

function listarSnapshotAtualParaReconciliacao_() {
  const mapa = {};
  let pageToken = null;
  do {
    const res = chamarApiComRetry_(() => Drive.Files.list({
      q: "'me' in owners and trashed = false",
      pageSize: 1000,
      pageToken: pageToken || undefined,
      fields: 'nextPageToken,files(id,mimeType,size,parents)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    }));
    const files = res.files || [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      mapa[f.id] = {
        id: f.id,
        mimeType: f.mimeType,
        size: f.size ? Number(f.size) : 0,
        parents: f.parents || []
      };
    }
    pageToken = res.nextPageToken || null;
  } while (pageToken);
  return mapa;
}

function etapa2_renomearOpcional() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const ss = abrirPlanilhaRelatorioValidada_();
    const abaCfg = ss.getSheetByName('Config');
    const mapaConfig = lerMapaConfig_(abaCfg);
    const modoDryRun = Boolean(CONFIG_EXECUCAO.DRY_RUN || String(mapaConfig['DRY_RUN']).toLowerCase() === 'true');

    if (modoDryRun) {
      Logger.log('DRY_RUN está ativo. Nenhuma renomeação foi executada.');
      return;
    }

    validarTravasDeProducaoPreVoo_(ss, mapaConfig);

    const abaRen = ss.getSheetByName('Sugestoes_Renomeacao');
    const abaWal = ss.getSheetByName('Log_Rollback');
    if (!abaRen || abaRen.getLastRow() <= 1) return;

    const dados = abaRen.getRange(2, 1, abaRen.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < dados.length; i++) {
      const [fileId, nomeAntigo, nomeNovo, , , aprovado] = dados[i];
      if (String(aprovado).trim().toUpperCase() !== 'SIM' || !fileId || !nomeNovo) continue;

      const logId = 'WAL_REN_' + Date.now() + '_' + i;
      const idxWal = gravarWalPendenteAntesDaAcao_(abaWal, logId, 'RENOMEACAO_OPCIONAL', 'RENOMEAR', fileId, nomeAntigo, nomeAntigo, nomeNovo);

      try {
        chamarApiComRetry_(() => Drive.Files.update({ name: String(nomeNovo) }, String(fileId), { supportsAllDrives: true, fields: 'id,name' }));
        atualizarStatusWal_(abaWal, idxWal, 'CONCLUIDO', 'Renomeado de [' + nomeAntigo + '] para [' + nomeNovo + ']');
        abaRen.getRange(i + 2, 6).setValue('CONCLUIDO');
      } catch (e) {
        atualizarStatusWal_(abaWal, idxWal, 'FALHA', e.message);
      }
    }
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function testes_logica_execucao() {
  Logger.log('Iniciando testes_logica_execucao() no Projeto B...');

  const snapOriginal = {
    id: 'FILE_1', parentId: 'FOLDER_A', md5Checksum: 'ABC', size: 1000,
    modifiedTime: '2026-11-10T10:00:00Z', ownedByMe: true, shared: false, driveId: '', canMove: true
  };
  const metaAoVivoIgual = {
    id: 'FILE_1', trashed: false, parents: ['FOLDER_A'], md5Checksum: 'ABC', size: 1000,
    modifiedTime: '2026-11-10T10:00:00Z', ownedByMe: true, shared: false, driveId: '',
    capabilities: { canMoveItemWithinDrive: true }
  };
  const metaAoVivoPaiMudou = {
    id: 'FILE_1', trashed: false, parents: ['FOLDER_OUTRA'], md5Checksum: 'ABC', size: 1000,
    modifiedTime: '2026-11-10T10:00:00Z', ownedByMe: true, shared: false, driveId: '',
    capabilities: { canMoveItemWithinDrive: true }
  };

  const resIgual = validarIntegridadePreMovimentoPuro_(snapOriginal, metaAoVivoIgual);
  assertExec_(resIgual.valido === true, 'Falha: Revalidação deveria aprovar arquivo inalterado.');

  const resMudou = validarIntegridadePreMovimentoPuro_(snapOriginal, metaAoVivoPaiMudou);
  assertExec_(resMudou.valido === false && resMudou.motivo.indexOf('Pai mudou') !== -1, 'Falha: Revalidação deveria bloquear arquivo que mudou de pasta.');

  const keeperEsperado = { id: 'KEEPER_1', md5Checksum: 'HASH_ORIGINAL', size: 5000 };
  const keeperVivoIntacto = { id: 'KEEPER_1', trashed: false, md5Checksum: 'HASH_ORIGINAL', size: 5000 };
  const flagSimulaLixeiraMock = Boolean(1);
  const keeperVivoNaLixeira = { id: 'KEEPER_1', trashed: flagSimulaLixeiraMock, md5Checksum: 'HASH_ORIGINAL', size: 5000 };

  assertExec_(validarIntegridadeDoMantidoPuro_(keeperEsperado, keeperVivoIntacto).valido === true, 'Falha ao validar Keeper intacto.');
  assertExec_(validarIntegridadeDoMantidoPuro_(keeperEsperado, keeperVivoNaLixeira).valido === false, 'Falha: Deveria bloquear quarentena se o Keeper estiver na lixeira.');

  let disparouDisjuntor = false;
  try {
    verificarDisjuntorDeErros_(5, 5, 10, 'Erro simulado de rede');
  } catch (e) {
    disparouDisjuntor = true;
  }
  assertExec_(disparouDisjuntor === true, 'Falha: Disjuntor deveria abortar com 5 erros consecutivos.');

  Logger.log('SUCESSO: Todos os testes em testes_logica_execucao() passaram!');
}

function assertExec_(condicao, msg) {
  if (!condicao) throw new Error('[ASSERTION FAILED] ' + msg);
}

function validarTravasDeProducaoPreVoo_(ss, mapaConfig) {
  const frasePlanilha = String(mapaConfig['FRASE_DE_CONFIRMACAO'] || '').trim();
  if (frasePlanilha !== CONFIG_EXECUCAO.FRASE_DE_CONFIRMACAO) {
    throw new Error(
      'TRAVA DE SEGURANÇA: Para executar com DRY_RUN = false, preencha na aba Config a chave FRASE_DE_CONFIRMACAO exatamente com: "' +
      CONFIG_EXECUCAO.FRASE_DE_CONFIRMACAO + '"'
    );
  }

  const tsAuditoria = Date.parse(String(mapaConfig['TIMESTAMP_ULTIMA_AUDITORIA'] || ''));
  const ignorarValidade24h = String(mapaConfig['IGNORAR_TRAVA_24H'] || '').trim().toUpperCase() === 'SIM';
  if (!ignorarValidade24h && (!tsAuditoria || (Date.now() - tsAuditoria) > CONFIG_EXECUCAO.VALIDADE_MAXIMA_AUDITORIA_HORAS * 3600 * 1000)) {
    throw new Error(
      'TRAVA DE SEGURANÇA: A auditoria tem mais de 24 horas (ou timestamp inválido). Refaça a auditoria no Projeto A para garantir metadados frescos.'
    );
  }

  const metaPlanilha = buscarMetadadosAoVivoArquivo_(ss.getId());
  if (!metaPlanilha || metaPlanilha.shared || !metaPlanilha.ownedByMe) {
    throw new Error('TRAVA DE PRIVACIDADE: A planilha Relatorio_Auditoria_Drive está compartilhada! Torne-a privada antes de executar.');
  }
}

function obterOuCriarPastaPrivadaPorCaminho_(caminhoCompleto, abaCfg, mapaConfig) {
  const chaveConfig = 'FOLDER_ID::' + caminhoCompleto;
  const idSalvo = String(mapaConfig[chaveConfig] || '').trim();

  if (idSalvo) {
    const metaPasta = buscarMetadadosAoVivoArquivo_(idSalvo);
    if (metaPasta && !metaPasta.trashed) {
      if (metaPasta.shared || !metaPasta.ownedByMe) {
        throw new Error('TRAVA DE PERMISSÃO: A pasta de destino (' + caminhoCompleto + ') está compartilhada! Abortando para não expor arquivos privados.');
      }
      return idSalvo;
    }
  }

  const partes = String(caminhoCompleto).replace(/^\/?Meu Drive\/?/i, '').split('/').filter(Boolean);
  let parentAtualId = CONFIG_EXECUCAO.PASTA_RAIZ_DO_ESCOPO || 'root';
  let caminhoAcumulado = '/Meu Drive';

  for (let i = 0; i < partes.length; i++) {
    const nomeSegmento = partes[i];
    caminhoAcumulado += '/' + nomeSegmento;
    const chaveSeg = 'FOLDER_ID::' + caminhoAcumulado;

    let idSegmento = String(mapaConfig[chaveSeg] || '').trim();
    let metaSeg = idSegmento ? buscarMetadadosAoVivoArquivo_(idSegmento) : null;

    if (!metaSeg || metaSeg.trashed) {
      const nomeEscapado = nomeSegmento.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
      const q = "name = '" + nomeEscapado + "' and '" + parentAtualId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and 'me' in owners and trashed = false";
      const busca = chamarApiComRetry_(() => Drive.Files.list({
        q: q,
        pageSize: 5,
        fields: 'files(id,name,shared,ownedByMe,trashed)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      }));

      if (busca.files && busca.files.length > 0) {
        metaSeg = busca.files[0];
      } else {
        const criada = chamarApiComRetry_(() => Drive.Files.create({
          name: nomeSegmento,
          mimeType: 'application/vnd.google-apps.folder',
          parents: [parentAtualId]
        }));
        metaSeg = buscarMetadadosAoVivoArquivo_(criada.id);
      }

      salvarChaveNaAbaConfig_(abaCfg, mapaConfig, chaveSeg, metaSeg.id, 'ID de pasta criada/mapeada sob demanda');
    }

    if (metaSeg.shared || !metaSeg.ownedByMe) {
      throw new Error('TRAVA DE SEGURANÇA: A pasta "' + caminhoAcumulado + '" está compartilhada. Nenhum arquivo será movido para ela.');
    }

    parentAtualId = metaSeg.id;
  }

  return parentAtualId;
}

function buscarMetadadosAoVivoArquivo_(fileId) {
  try {
    return chamarApiComRetry_(() => Drive.Files.get(String(fileId), {
      fields: 'id,name,mimeType,parents,trashed,md5Checksum,size,modifiedTime,ownedByMe,shared,driveId,capabilities(canMoveItemWithinDrive)',
      supportsAllDrives: true
    }));
  } catch (e) {
    return null;
  }
}

function pastaExisteEAcessivel_(folderId) {
  if (!folderId || folderId === 'root') return true;
  const meta = buscarMetadadosAoVivoArquivo_(folderId);
  return Boolean(meta && !meta.trashed && meta.ownedByMe && !meta.shared);
}

function moverArquivoApiV3ComRetry_(fileId, oldParentId, newParentId) {
  const resultado = chamarApiComRetry_(() => Drive.Files.update(
    {},
    String(fileId),
    {
      addParents: String(newParentId),
      removeParents: String(oldParentId),
      supportsAllDrives: true,
      fields: 'id,parents'
    }
  ));

  if (!resultado || !resultado.parents || resultado.parents.indexOf(String(newParentId)) === -1) {
    throw new Error('A API não confirmou o novo parentId (' + newParentId + ') para o arquivo ' + fileId);
  }
  return resultado;
}

function gravarWalPendenteAntesDaAcao_(abaWal, logId, onda, tipoOp, fileId, nome, paiOrigem, paiDestino) {
  const novaLinha = abaWal.getLastRow() + 1;
  abaWal.getRange(novaLinha, 1, 1, 10).setValues([[
    logId,
    onda,
    tipoOp,
    fileId,
    nome,
    paiOrigem,
    paiDestino,
    new Date().toISOString(),
    'PENDENTE',
    'Write-Ahead Log gravado antes da chamada de API'
  ]]);
  SpreadsheetApp.flush();
  return novaLinha;
}

function atualizarStatusWal_(abaWal, linhaIndex, novoStatus, detalhes) {
  abaWal.getRange(linhaIndex, 9, 1, 2).setValues([[novoStatus, detalhes]]);
}

function registrarLinhaWalCompleta_(abaWal, onda, tipoOp, fileId, nome, paiOrigem, paiDestino, status, detalhes) {
  const novaLinha = abaWal.getLastRow() + 1;
  abaWal.getRange(novaLinha, 1, 1, 10).setValues([[
    'WAL_SKIP_' + Date.now(),
    onda,
    tipoOp,
    fileId,
    nome,
    paiOrigem,
    paiDestino,
    new Date().toISOString(),
    status,
    detalhes
  ]]);
}

function contarItensJaProcessadosNaOnda_(abaWal, nomeOnda) {
  if (!abaWal || abaWal.getLastRow() <= 1) return 0;
  const dados = abaWal.getRange(2, 1, abaWal.getLastRow() - 1, 9).getValues();
  let count = 0;
  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][1]) === nomeOnda && String(dados[i][8]) === 'CONCLUIDO') {
      count++;
    }
  }
  return count;
}

function abrirPlanilhaRelatorioValidada_() {
  if (CONFIG_EXECUCAO.SPREADSHEET_ID_RELATORIO) {
    return SpreadsheetApp.openById(CONFIG_EXECUCAO.SPREADSHEET_ID_RELATORIO);
  }
  const q = "name = 'Relatorio_Auditoria_Drive' and mimeType = 'application/vnd.google-apps.spreadsheet' and 'me' in owners and trashed = false";
  const busca = chamarApiComRetry_(() => Drive.Files.list({ q: q, pageSize: 1, fields: 'files(id,name)' }));
  if (!busca.files || busca.files.length === 0) {
    throw new Error('Planilha Relatorio_Auditoria_Drive não encontrada. Preencha SPREADSHEET_ID_RELATORIO.');
  }
  return SpreadsheetApp.openById(busca.files[0].id);
}

function inicializarAbasProjetoB_(ss) {
  const esquemasB = [
    {
      nome: 'Simulacao',
      colunas: ['Timestamp (ISO)', 'Onda Prevista', 'Tipo Operação', 'FileID', 'Arquivo', 'Simulação ([Arquivo] -> [Ação] -> [Destino ou motivo])', 'Status Validação']
    },
    {
      nome: 'Log_Rollback',
      colunas: ['LogID', 'Lote / Onda', 'Tipo Operação', 'FileID', 'Nome do Arquivo', 'ParentID_Origem', 'ParentID_Destino', 'Timestamp', 'Status (PENDENTE / CONCLUIDO / FALHA / REVERTIDO)', 'Detalhes']
    },
    {
      nome: 'Reconciliacao',
      colunas: ['Métrica de Integridade', 'Valor Antes (Inventário)', 'Valor Depois (Drive Ao Vivo)', 'Status da Reconciliação']
    }
  ];

  esquemasB.forEach(esq => {
    let aba = ss.getSheetByName(esq.nome);
    if (!aba) {
      aba = ss.insertSheet(esq.nome);
    }
    if (aba.getLastRow() === 0) {
      aba.getRange(1, 1, 1, esq.colunas.length).setValues([esq.colunas]).setFontWeight('bold');
    }
  });

  const abaCfg = ss.getSheetByName('Config');
  const mapaCfg = lerMapaConfig_(abaCfg);
  if (!mapaCfg['ESTADO_ONDA_ATUAL']) {
    salvarChaveNaAbaConfig_(abaCfg, mapaCfg, 'ESTADO_ONDA_ATUAL', 'PILOTO_20', 'Controla o portão de ondas (PILOTO_20 -> ONDA_200 -> ONDA_FINAL)');
  }
  if (!mapaCfg['PILOTO_CONFERIDO']) {
    salvarChaveNaAbaConfig_(abaCfg, mapaCfg, 'PILOTO_CONFERIDO', 'NAO', 'Mude para SIM após conferir os 20 itens do lote piloto no Drive');
  }
  if (!mapaCfg['ONDA_2_CONFERIDA']) {
    salvarChaveNaAbaConfig_(abaCfg, mapaCfg, 'ONDA_2_CONFERIDA', 'NAO', 'Mude para SIM após conferir os 200 itens da Onda 2 no Drive');
  }
}

function lerMapaConfig_(abaCfg) {
  const mapa = {};
  if (!abaCfg || abaCfg.getLastRow() <= 1) return mapa;
  const dados = abaCfg.getRange(2, 1, abaCfg.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < dados.length; i++) {
    const k = String(dados[i][0] || '').trim();
    if (k) mapa[k] = String(dados[i][1] || '').trim();
  }
  return mapa;
}

function salvarChaveNaAbaConfig_(abaCfg, mapaMemoria, chave, valor, descricao) {
  mapaMemoria[chave] = String(valor);
  const ultima = abaCfg.getLastRow();
  if (ultima > 1) {
    const dados = abaCfg.getRange(2, 1, ultima - 1, 1).getValues();
    for (let i = 0; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === chave) {
        abaCfg.getRange(i + 2, 2, 1, 2).setValues([[String(valor), descricao || '']]);
        SpreadsheetApp.flush();
        return;
      }
    }
  }
  abaCfg.getRange(ultima + 1, 1, 1, 3).setValues([[chave, String(valor), descricao || '']]);
  SpreadsheetApp.flush();
}

function regravaMatrizEmLotes_(aba, linhaInicial, colInicial, matriz) {
  if (!matriz || matriz.length === 0) return;
  const numCols = matriz[0].length;
  const lote = 2000;
  for (let i = 0; i < matriz.length; i += lote) {
    const fatia = matriz.slice(i, i + lote);
    aba.getRange(linhaInicial + i, colInicial, fatia.length, numCols).setValues(fatia);
  }
}

function limparDadosMantendoCabecalho_(aba) {
  if (!aba) return;
  const u = aba.getLastRow();
  if (u > 1) {
    aba.getRange(2, 1, u - 1, aba.getLastColumn()).clearContent();
  }
}

function chamarApiComRetry_(fn) {
  const maxTentativas = 5;
  for (let t = 1; t <= maxTentativas; t++) {
    try {
      return fn();
    } catch (e) {
      const msg = String(e && e.message ? e.message : e);
      const transiente = /(429|500|503|rateLimitExceeded|userRateLimitExceeded|quotaExceeded|backendError)/i.test(msg);
      if (!transiente || t === maxTentativas) throw e;
      Utilities.sleep(Math.pow(2, t) * 1000 + Math.floor(Math.random() * 500));
    }
  }
}

function agendarGatilhoProjetoB_(nomeFuncao) {
  limparGatilhosProjetoB_(nomeFuncao);
  ScriptApp.newTrigger(nomeFuncao).timeBased().after(60 * 1000).create();
}

function limparGatilhosProjetoB_(nomeFuncao) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (!nomeFuncao || triggers[i].getHandlerFunction() === nomeFuncao) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function registrarLogGeral_(ss, funcao, idAlvo, acao, resultado) {
  try {
    const abaLog = ss.getSheetByName('Log');
    if (abaLog) {
      abaLog.getRange(abaLog.getLastRow() + 1, 1, 1, 5).setValues([[
        new Date().toISOString(), funcao, idAlvo, acao, resultado
      ]]);
    }
  } catch (e) {
    Logger.log('Falha ao gravar log: ' + e.message);
  }
}
