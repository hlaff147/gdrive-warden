# gdrive-warden

Zero-loss Google Drive auditor, cryptographic MD5/SHA-256 deduplicator, and taxonomy organizer built for **Google Drive API v3** (Google Apps Script + CLI workflows), featuring read-only isolation, write-ahead logging (WAL), circuit breaker, pre-move revalidation, and idempotent rollback.

## Core Safety Guarantees

1. **Zero Deletion by Construction**: Never calls `delete`, `remove`, `emptyTrash`, or `setTrashed`. Approved Level 1 exact duplicates (`MD5` + `SHA-256` + `size`) are isolated exclusively into `/Meu Drive/_QUARENTENA_DUPLICATAS_PARA_REVISAO` while preserving a deterministic Keeper.
2. **Privilege Separation**:
   - **Project A (`drive.readonly`)**: Audits inventory, detects protected repositories/projects, runs 3-level deduplication, and generates routing plans without write permissions.
   - **Project B (`drive`)**: Executes approved moves in gated waves (`Pilot 20` -> `Wave 200` -> `Final`), backed by a Write-Ahead Log (`Log_Rollback`), live pre-move revalidation, reconciliation (`bytes` + `file count`), and idempotent rollback.
3. **Permission & Repository Preservation**: Automatically skips shared files, shared folders, non-owned items, shortcuts, Shared Drives, and internal trees of code repositories (`.git`, `package.json`, `pyproject.toml`, `pom.xml`, `.vs`, etc.).
4. **EXIF-Strict Media Organization**: Routes photos into `05_Midia_e_Fotos/Anos/YYYY/YYYY-MM` strictly when capture EXIF (`imageMediaMetadata.time`) is present; photos lacking EXIF go to `05_Midia_e_Fotos/Sem_Data_Confiavel` rather than misusing upload timestamps (`createdTime`).

## Repository Structure

```text
.
├── README.md
├── prompts/
│   └── mega-prompt-ptbr.md                  # Engineering specification (PT-BR)
├── src/
│   ├── project-a-audit-readonly/
│   │   ├── appsscript.json                  # Read-only OAuth manifest (drive.readonly)
│   │   ├── Codigo.gs                        # Stage 2: Paginated audit & 3-tier deduplication
│   │   └── Roteamento_Etapa3.gs             # Stage 3: Confidence-scored taxonomy routing
│   └── project-b-execution-rollback/
│       ├── appsscript.json                  # Execution OAuth manifest
│       ├── ProjetoB_Execucao_Rollback.gs    # Stage 4: WAL, waves, circuit breaker, rollback
│       └── Gerador_Teste.gs                 # Synthetic test harness (_TESTE_ORGANIZACAO)
├── cli/
│   └── rollback_wal.py                      # Local WAL inspection & rollback helper
└── docs/
    └── playbook-homologacao.md              # End-to-end homologation & operations playbook
```

## Static Verification & Unit Tests

Run the local verification suite (forbidden-API static scanner + synthetic logic assertions for deduplication, routing invariants, pre-move revalidation, and circuit breaker):

```bash
node tests/run_unit_tests.js
```
