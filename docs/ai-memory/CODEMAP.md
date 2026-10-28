# Code map

Last audited against 4c203a9420c971827e1f3559d77b362ff9803212 on 2026-10-28.

| Area | Path | Purpose |
| --- | --- | --- |
| Import entry point | IMPORTAR_UNIALIADOS_OPTIMIZADO.cmd | Windows entry point for the import process |
| Import implementation | Importar_Unialiados_Optimizado.ps1 | PowerShell import and setup logic |
| Workflow definition | Unialiados - Prospección local (v17.4 fix gramatical).json | n8n workflow definition; inspect nodes and credential references carefully |
| Operational workbook | Unialiados_Operacion_Unificada.xlsm | Excel macro-enabled operational artifact; do not expose contents in public notes |

## Selective lookup

- Import behavior: inspect the CMD launcher and the PowerShell script together.
- Workflow behavior: inspect the JSON node list, connections and configuration references without copying credential values.
- Operational data: open the workbook only when the task explicitly concerns it; record structure and verified outcomes, not customer rows or secrets.
- Verification: use a PowerShell dry run or controlled test environment where available. This repository has no package manifest or automated test command in the audited baseline.
