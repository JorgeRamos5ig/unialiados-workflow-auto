# Technical decisions

## D-001 — Memory is a map, not a source copy

The scripts, workflow JSON and workbook are authoritative. Memory records paths, behavior, constraints and evidence without duplicating code or operational data.

## D-002 — Credentials stay out of memory

Do not record n8n credential values, tokens, connection strings, customer data, workbook rows or macro payloads. Refer to the relevant file or secure environment only.

## D-003 — Validate automation changes safely

Changes to import paths, node IDs, workflow connections or macros require a controlled validation before they are described as working.

## D-004 — Public Obsidian stays non-sensitive

The public dashboard may link to this private repository and show compact orientation only. It must not contain source, exports, credentials or operational records.

## D-005 — Campaign rotation per contact (2026-10-28, explicit request, overrides earlier anti-alternation rule)

As of commit 9c31e17 on master (v17.5), each lead rotates through all 3 campaigns (Unialiados, Unisoft, AXA Colpatria) across its own 1st/2nd/3rd contact, instead of being locked to a single campaign forever (the old `campaignLockDays` behavior). This was an explicit instruction from Jorge on 2026-10-28, and it intentionally contradicts the project's earlier documented rule ("never alternate campaigns to the same contact within a short period; keep one campaign per contact for ~2 months"). That earlier rule is not wrong as a general spam-avoidance principle — it was simply superseded for this specific mechanic by a direct decision.

Controlled by `campaignRotationEnabled` in Config (default `true`). Setting it to `false` restores the old lock-forever behavior exactly. A campaign manually set by staff in the "Campaña asignada" Excel column is never touched by rotation — only automatically-decided campaigns rotate.

Touched nodes: `Asignar Campaña (Reglas)`, `Validar Campaña Asignada` (the actual authoritative lock check — fixing only the first node would not have been enough), `Preparar Archivo de Actualización Local` (must reset "Campaña asignada" back to "Auto" after an automatic send, or the next contact would see a concrete value and treat it as a manual override).

## D-006 — AXA email goal is to book a meeting, not to resolve the pitch by email (2026-10-28, audio from Lorena)

Lorena's audio on 2026-10-28 clarified the AXA Colpatria campaign's commercial mechanics: Unialiados is a broker/intermediary that earns a commission from AXA for bringing in companies; the email's only job is to get the prospect to agree to a short meeting, and Unialiados' commercial team then visits and explains in person. This is consistent with (not a reversal of) the "exploratory conversation" framing already built from the 15-oct meeting with Cristian Pardo (AXA) — it sharpens the CTA to explicitly request a meeting with the commercial team, and adds an explicit rule that the email must never mention Unialiados' commission. See `Construir Prompt de Correo`.

## D-007 - Validation fixes for language leaks and duplicated CTA (2026-10-28)

The workflow JSON was patched locally after two production-like validation findings: the Portuguese detector incorrectly treated normal Spanish words such as esta/está and tambien as Portuguese, while the English detector missed the bare English word in; generated output could also repeat the CTA question inside paragraphs. The fix keeps only unambiguous Portuguese spellings, expands the English leak list, adds ctaDuplicatedInParagraphs, and enforces exactly two paragraphs for FIRST/SECOND contacts. Construir Prompt de Correo now returns targeted retry hints for the paragraph-count and CTA-duplication errors.

Evidence checked: the extracted Preparar Datos de Envio JavaScript compiles with new Function; the lead-202 fixture is rejected for the English leak, and the known lead-203 fixture no longer triggers the Portuguese detector (it still correctly rejects the fixture when its first paragraph repeats the system-generated greeting). The source-of-truth workflow JSON remains the file named in CODEMAP; these changes were observed in the working tree and are not claimed as committed to master by this memory entry.
