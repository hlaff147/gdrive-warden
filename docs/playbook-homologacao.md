# Playbook de Homologação e Operação Fim a Fim

Este guia descreve a sequência operacional segura para homologar e operar a arquitetura Zero-Loss do `gdrive-warden` via Google Apps Script (Projetos A e B) ou via CLI local.

---

## Pré-Requisitos Obrigatórios antes da Execução

1. **Auditoria Recente**: menos de 24 horas desde a geração do inventário.
2. **Backup do Relatório**: cópia de segurança da planilha `Relatorio_Auditoria_Drive`.
3. **Backup de Documentos Críticos**: exportação preventiva (Google Takeout) para itens sensíveis.
4. **Drive para Desktop Pausado**: pause a sincronização local durante os lotes de movimentação.
5. **Congelamento de Edições**: não mova nem edite arquivos manualmente durante as ondas.
6. **Privacidade da Planilha**: confirme que `Relatorio_Auditoria_Drive` está privada (`shared = false`).

---

## Fase 1 — Homologação na Pasta Sintética (`_TESTE_ORGANIZACAO`)

1. No **Projeto B**, execute `etapa0_criarAmbienteTesteSintetico()` ([Gerador_Teste.gs](../src/project-b-execution-rollback/Gerador_Teste.gs)) e copie o ID gerado nos logs.
2. No **Projeto A** ([Codigo.gs](../src/project-a-audit-readonly/Codigo.gs) e [Roteamento_Etapa3.gs](../src/project-a-audit-readonly/Roteamento_Etapa3.gs)):
   - Cole o ID em `PASTA_RAIZ_DO_ESCOPO` e remova temporariamente `'_TESTE_ORGANIZACAO'` de `PASTAS_PROTEGIDAS_NOMES`.
   - Execute `testes_logica()` e confirme `SUCESSO`.
   - Execute `etapa1_auditoria()` e aguarde `CONCLUIDO` na aba `Log`.
   - Execute `testes_logica_roteamento()` e confirme `SUCESSO`.
   - Execute `etapa3_gerarPlanoOrganizacao()` e aguarde `CONCLUIDO`.
3. Na planilha `Relatorio_Auditoria_Drive`, revise as abas `Duplicatas_Detectadas` (apenas `NIVEL_1_EXATA` onde `Mantido? = NAO`) e `Plano_Organizacao`, marcando `SIM` na coluna `Aprovado`.
4. No **Projeto B** ([ProjetoB_Execucao_Rollback.gs](../src/project-b-execution-rollback/ProjetoB_Execucao_Rollback.gs)):
   - Cole o ID da planilha em `SPREADSHEET_ID_RELATORIO` e o ID de `_TESTE_ORGANIZACAO` em `PASTA_RAIZ_DO_ESCOPO`.
   - Execute `testes_logica_execucao()` e confirme `SUCESSO`.
   - Com `DRY_RUN: true`, execute `etapa2_executar()` e valide a aba `Simulacao`.
   - Preencha `FRASE_DE_CONFIRMACAO` na aba `Config` com `EU REVISEI A PLANILHA E AUTORIZO A EXECUCAO`, ajuste `DRY_RUN` para `false` e execute `etapa2_executar()`.
   - Valide a aba `Reconciliacao` (`INTEGRIDADE 100%`) e teste `etapa3_rollback()` para confirmar o retorno idempotente à pasta original.

---

## Fase 2 — Execução em Produção com Portões de Onda

1. Volte `PASTA_RAIZ_DO_ESCOPO: ''` e recoloque `'_TESTE_ORGANIZACAO'` em `PASTAS_PROTEGIDAS_NOMES`.
2. Execute `etapa1_auditoria()` e `etapa3_gerarPlanoOrganizacao()` no Projeto A (somente leitura).
3. Revise por amostragem (mínimo de 20 linhas por categoria, além de todos os itens de Nível 2 e Nível 3) e aprove linha a linha (`SIM`).
4. No Projeto B:
   - **Simulação**: rode `etapa2_executar()` com `DRY_RUN: true`.
   - **Onda 1 (Piloto — 20 itens)**: com `DRY_RUN: false`, rode `etapa2_executar()`. Confira no Drive e marque `PILOTO_CONFERIDO = SIM` na aba `Config`.
   - **Onda 2 (200 itens)**: rode `etapa2_executar()`. Confira no Drive e marque `ONDA_2_CONFERIDA = SIM` na aba `Config`.
   - **Onda Final**: rode `etapa2_executar()` até concluir a fila e valide a aba `Reconciliacao`.

---

## Fase 3 — Quarentena de 30 Dias e Modos de Falha

- Mantenha `_QUARENTENA_DUPLICATAS_PARA_REVISAO` intacta por pelo menos **30 dias** antes de qualquer limpeza manual.
- **Modos de falha cobertos**:
  - *Estouro de tempo (~6 min)*: pausa automática aos 4,5 min com cursor em `PropertiesService` e gatilho temporário autolimpante.
  - *Limite de taxa (429 / 500 / 503)*: retry com backoff exponencial + jitter.
  - *Alteração concorrente*: revalidação pré-movimento (`PULADO_MUDOU`) se pai, tamanho, MD5, `modifiedTime` ou compartilhamento mudarem.
  - *Pasta original removida antes do rollback*: restauração automática para `/Meu Drive/_RESTAURADOS`.
  - *Cascata de falhas*: disjuntor (circuit breaker) aborta após 5 erros consecutivos ou taxa de erro > 5%.
