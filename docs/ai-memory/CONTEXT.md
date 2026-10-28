# Technical memory: Unialiados workflow automation

> Compact orientation for Claude Code and Codex. Repository code and workflow files are authoritative.

## Identity

- Repository: https://github.com/JorgeRamos5ig/unialiados-workflow-auto
- Default branch: master
- Current audited commit: 4c203a9420c971827e1f3559d77b362ff9803212 (2026-10-28)
- Stack: PowerShell 96%, Batchfile 4%; includes an n8n workflow JSON and an XLSM operational workbook.
- Visibility: private. Never copy workflow credentials, customer data, workbook contents or secrets into the public Obsidian hub.

## Current reality

- The repository is an automation package for importing and operating the Unialiados workflow, not a conventional application with a package manifest or test suite.
- The tracked files are IMPORTAR_UNIALIADOS_OPTIMIZADO.cmd, Importar_Unialiados_Optimizado.ps1, the Unialiados Prospección local workflow JSON, and Unialiados_Operacion_Unificada.xlsm.
- Verify the PowerShell and workflow JSON before changing import paths, node IDs, credentials, workbook macros or execution assumptions.

## Working rules

1. Read this file and CODEMAP.md before working.
2. Read DECISIONS.md for safety, credentials, data-handling or workflow-contract changes.
3. Read HANDOFF.md when continuing prior work.
4. Inspect only the relevant script/workflow sections and verify changes with a safe dry run when available.
5. Update memory incrementally with evidence; never record secrets, tokens, customer data or unverified claims.
