/**
 * ============================================================================
 * GERADOR DE MASSA DE TESTE SINTÉTICA (_TESTE_ORGANIZACAO)
 * ============================================================================
 */

function etapa0_criarAmbienteTesteSintetico() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Execução abortada: outra instância está rodando.');
    return;
  }

  try {
    const nomePastaTeste = '_TESTE_ORGANIZACAO';
    const iterador = DriveApp.getRootFolder().getFoldersByName(nomePastaTeste);

    if (iterador.hasNext()) {
      const existente = iterador.next();
      Logger.log('ABORTADO POR IDEMPOTÊNCIA: A pasta ' + nomePastaTeste + ' já existe.');
      Logger.log('PASTA_RAIZ_DO_ESCOPO = "' + existente.getId() + '"');
      return;
    }

    const pastaRaizTeste = DriveApp.getRootFolder().createFolder(nomePastaTeste);
    const idRaizTeste = pastaRaizTeste.getId();

    const pastaProjeto = pastaRaizTeste.createFolder('meu_microservico_teste');
    pastaProjeto.createFile('package.json', '{"name": "microservico-teste", "version": "1.0.0"}', MimeType.PLAIN_TEXT);
    pastaProjeto.createFile('README.md', '# Projeto de Teste Blindado\nConteúdo interno nunca deve ser movido.', MimeType.PLAIN_TEXT);

    const pastaOrganizada = pastaRaizTeste.createFolder('Documentos_Fiscais_Salvos');

    const conteudoDuplicataA = 'CONTEUDO_BINARIO_IDENTICO_GRUPO_A_IRPF_2026_HASH_TESTE_001';
    pastaOrganizada.createFile('Comprovante_IRPF_2026.txt', conteudoDuplicataA, MimeType.PLAIN_TEXT);
    pastaRaizTeste.createFile('Comprovante_IRPF_2026 (1).txt', conteudoDuplicataA, MimeType.PLAIN_TEXT);
    pastaRaizTeste.createFile('Cópia de Comprovante_IRPF_2026.txt', conteudoDuplicataA, MimeType.PLAIN_TEXT);

    const conteudoDuplicataB = 'CONTEUDO_BINARIO_IDENTICO_GRUPO_B_BOLETO_HASH_TESTE_002';
    pastaRaizTeste.createFile('Boleto_Condominio.txt', conteudoDuplicataB, MimeType.PLAIN_TEXT);
    pastaRaizTeste.createFile('Boleto_Condominio (1).txt', conteudoDuplicataB, MimeType.PLAIN_TEXT);

    pastaRaizTeste.createFile('Contrato_Locacao_Assinado.txt', 'CLAUSULA_1_VALOR_ALUGUEL_1500_VERSAO_A', MimeType.PLAIN_TEXT);
    pastaRaizTeste.createFile('Contrato_Locacao_Assinado (1).txt', 'CLAUSULA_1_VALOR_ALUGUEL_1800_VERSAO_B', MimeType.PLAIN_TEXT);

    const arquivosAltaConfianca = [
      { nome: 'CNH_Digital_Identidade.pdf', texto: 'Mock PDF CNH Identidade Civil', mime: MimeType.PDF },
      { nome: 'Declaracao_IRPF_Recibo_2025.pdf', texto: 'Mock PDF Imposto de Renda', mime: MimeType.PDF },
      { nome: 'Lista_de_Compras_Casa_Nova_Geladeira.txt', texto: 'Mock Lista Eletrodomésticos', mime: MimeType.PLAIN_TEXT },
      { nome: 'Resultado_Hemograma_Exame_Laboratorial.pdf', texto: 'Mock Exame de Saúde', mime: MimeType.PDF },
      { nome: 'Curriculo_Portfolio_Engenheiro_Software.pdf', texto: 'Mock Currículo', mime: MimeType.PDF },
      { nome: 'Certificado_Conclusao_Curso_Cloud.pdf', texto: 'Mock Certificado', mime: MimeType.PDF },
      { nome: 'Holerite_Contracheque_Mensal.pdf', texto: 'Mock Financeiro Profissional', mime: MimeType.PDF },
      { nome: 'Dissertacao_Mestrado_Circuit_Breaker_UFPE.pdf', texto: 'Mock Artigo Acadêmico', mime: MimeType.PDF },
      { nome: 'Apostila_Ebook_Arquitetura_Sistemas.pdf', texto: 'Mock Livro/PDF Estudo', mime: MimeType.PDF },
      { nome: 'Lista_Exercicios_Ementa_Disciplina.pdf', texto: 'Mock Disciplina Acadêmica', mime: MimeType.PDF },
      { nome: 'Comparativo_Braco_Articulado_F80N_Monitor.txt', texto: 'Mock Hardware e Setup', mime: MimeType.PLAIN_TEXT },
      { nome: 'argocd_config_deployment.yaml', texto: 'apiVersion: argoproj.io/v1alpha1', mime: MimeType.PLAIN_TEXT },
      { nome: 'automacao_backup.py', texto: 'print("Teste script Python solto")', mime: MimeType.PLAIN_TEXT },
      { nome: 'metricas_experimento.csv', texto: 'id,latencia_ms\n1,12.4\n2,15.1', mime: MimeType.CSV }
    ];

    for (let i = 0; i < arquivosAltaConfianca.length; i++) {
      const item = arquivosAltaConfianca[i];
      pastaRaizTeste.createFile(item.nome, item.texto, item.mime);
    }

    pastaRaizTeste.createFile('recibo_avulso.pdf', 'Mock palavra isolada', MimeType.PDF);
    pastaRaizTeste.createFile('boleto_taxa_inscricao_ufpe.pdf', 'Mock ambiguidade entre categorias', MimeType.PDF);
    const png1x1Base64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    const blobPngSemExif = Utilities.newBlob(Utilities.base64Decode(png1x1Base64), 'image/png', 'WhatsApp Image 2026-11-15 at 10.00.00.png');
    pastaRaizTeste.createFile(blobPngSemExif);

    pastaRaizTeste.createFile('scan00239.pdf', 'Mock arquivo escaneado sem nome descritivo', MimeType.PDF);
    pastaRaizTeste.createFile('arquivo_desconhecido_123.dat', 'Dados sem padrão claro', MimeType.PLAIN_TEXT);

    const docNativo = DocumentApp.create('Crie um docs completo com tudo isso por favor - Teste');
    const fileDocNativo = DriveApp.getFileById(docNativo.getId());
    fileDocNativo.moveTo(pastaRaizTeste);

    Logger.log('================================================================');
    Logger.log('AMBIENTE SINTÉTICO CRIADO COM SUCESSO!');
    Logger.log('Nome da Pasta: ' + nomePastaTeste);
    Logger.log('Copie este ID para PASTA_RAIZ_DO_ESCOPO nos Projetos A e B:');
    Logger.log(idRaizTeste);
    Logger.log('IMPORTANTE: Ao testar esta pasta no Projeto A, remova temporariamente');
    Logger.log('"_TESTE_ORGANIZACAO" da lista PASTAS_PROTEGIDAS_NOMES.');
    Logger.log('================================================================');

  } finally {
    lock.releaseLock();
  }
}
