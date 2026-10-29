# Technical memory: Unialiados workflow automation

> Compact orientation for Claude Code and Codex. Repository code and workflow files are authoritative.

## Identity

- Repository: https://github.com/JorgeRamos5ig/unialiados-workflow-auto
- Default branch: master
- Current audited commit: 9c31e17 on master (2026-10-28, v17.5 — campaign rotation per contact, AXA CTA aligned to Lorena's audio; see DECISIONS.md D-005/D-006)
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
- Latest published master commit: 5927c9f on master (2026-10-28, v17.6). The memory branch is documentation-only; see DECISIONS.md D-005 through D-007 for the latest workflow decisions and validation findings.

## Latest validation handoff

- The local workflow JSON contains fixes for false Portuguese positives, a missed English `in`, duplicated CTA text, and extra CTA-as-paragraph output.
- FIRST and SECOND contacts now require exactly two `paragraphs`; the CTA remains only in `cta`.
- The validation code passed a JavaScript syntax check. This does not replace a controlled n8n import/run or a real Ollama/Outlook/Excel integration test.
- Local XLSM changes, backup files and generated operational material are intentionally excluded from this memory branch.
