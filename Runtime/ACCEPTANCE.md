# Acceptance test — 2026-10-29

Command:

```text
node Runtime/acceptance_test.js
```

Environment verified:

- Ollama API: `http://127.0.0.1:11434/api/chat`
- Model: `qwen3.5:4b`
- Workflow JSON: `Unialiados - Prospección local (v17.4 fix gramatical).json`

Matrix:

| Campaign | FIRST | SECOND | THIRD |
| --- | --- | --- | --- |
| Unialiados | PASS | PASS | PASS |
| AXA Colpatria | PASS | PASS | PASS |
| Unisoft | PASS | PASS | PASS |

Final result: **9/9 accepted, 0 rejected, 0 advertising-signal matches**.

Campaign rotation test:

```text
node Runtime/rotation_acceptance_test.js
```

Result: **PASS** — the same lead rotated across `FIRST → SECOND → THIRD` as `Unialiados → AXA Colpatria → Unisoft`, with every stage accepted by the authoritative campaign validator.

The harness exercises the embedded `Construir Prompt de Correo` and `Preparar Datos de Envío` code and uses Ollama responses, including the configured retry path. It does not send mail, write Excel, call Tavily, or execute the full n8n runtime. SMTP delivery, workbook persistence, and end-to-end n8n execution still require a controlled integration test with a test mailbox and a workbook backup.
