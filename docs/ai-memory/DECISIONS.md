# Technical decisions

## D-001 — Memory is a map, not a source copy

The scripts, workflow JSON and workbook are authoritative. Memory records paths, behavior, constraints and evidence without duplicating code or operational data.

## D-002 — Credentials stay out of memory

Do not record n8n credential values, tokens, connection strings, customer data, workbook rows or macro payloads. Refer to the relevant file or secure environment only.

## D-003 — Validate automation changes safely

Changes to import paths, node IDs, workflow connections or macros require a controlled validation before they are described as working.

## D-004 — Public Obsidian stays non-sensitive

The public dashboard may link to this private repository and show compact orientation only. It must not contain source, exports, credentials or operational records.
