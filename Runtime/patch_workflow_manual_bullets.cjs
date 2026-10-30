const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', 'Unialiados - Prospección local (v17.4 fix gramatical).json');
const workflow = JSON.parse(fs.readFileSync(file, 'utf8'));
const node = (name) => workflow.nodes.find((n) => n.name === name);
const replaceExactly = (source, oldText, newText, label) => {
  const count = source.split(oldText).length - 1;
  if (source.includes(newText)) return source;
  if (count !== 1) throw new Error(`${label}: expected one match, found ${count}`);
  return source.replace(oldText, newText);
};
const setAssignment = (name, value) => {
  const assignments = node('Config: batch y remitente').parameters.assignments.assignments;
  let item = assignments.find((x) => x.name === name);
  if (!item) {
    const id = `manual-${name}`;
    const type = typeof value === 'boolean' ? 'boolean' : typeof value === 'number' ? 'number' : 'string';
    item = { id, name, value, type };
    assignments.push(item);
  } else item.value = value;
};

// La reunión acordó que el cliente filtra el Excel, marca a quién contactará y pulsa Iniciar.
setAssignment('sendEnabled', false);
setAssignment('maxEmailsPerDay', 50);
setAssignment('maxLeadsPerRun', 50);
setAssignment('senderEmail', 'PENDIENTE_CONFIGURAR');
setAssignment('workbookPath', 'C:\\Automatizacion\\Unialiados\\Unialiados_Operacion_Unificada.xlsm');
setAssignment('excelHelperPath', 'C:\\Automatizacion\\Unialiados\\UnialiadosExcelLocal.ps1');
setAssignment('runtimeDirectory', 'C:\\Automatizacion\\Unialiados\\Runtime');
setAssignment('backupDirectory', 'C:\\Automatizacion\\Unialiados\\Respaldos');
setAssignment('campaignAutoEnabled', false);
setAssignment('campaignRotationEnabled', false);
setAssignment('aiCampaignMode', 'off');
setAssignment('axaAutoAssignEnabled', false);
setAssignment('enforceOneEmailPerMailbox', true);
setAssignment('manualBatchControlEnabled', true);
setAssignment('senderIdentityConfirmed', false);
setAssignment('sendLimitsConfirmed', false);
setAssignment('followUpTimingConfirmed', false);
setAssignment('smtpCredentialConfigured', false);
setAssignment('optOutProcessVerified', false);
setAssignment('axaExclusionsComplete', false);
setAssignment('axaCommercialApproved', true);
setAssignment('axaExcludedPattern', 'agro|agricultur|ganader|mineri|socavon|siembra|cultivo|temporal|servicio domestico|aseo|limpieza');

// Desactiva el cron: la selección explícita en la hoja y el arranque manual son obligatorios.
const schedule = node('Programado (según cron)');
if (!schedule) throw new Error('Schedule trigger not found');
schedule.disabled = true;

// Añade una prevalidación que hace visibles los pendientes si alguien intenta habilitar envíos.
const configNode = node('Config: batch y remitente');
const readCompany = node('Leer Empresas y Archivos');
if (!configNode || !readCompany) throw new Error('Could not locate configuration/read nodes');
const preflightName = 'Validar configuración para envío real';
if (!node(preflightName)) {
  const preflight = {
    parameters: {
      jsCode: `const config = $input.first().json;\nif (config.sendEnabled === true) {\n  const issues = [];\n  const email = String(config.senderEmail || '').trim();\n  if (!/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}$/i.test(email) || /pendiente|dominio\\.com/i.test(email)) issues.push('configura el correo remitente real');\n  if (config.senderIdentityConfirmed !== true) issues.push('confirma con el cliente el nombre y el buzón remitente');\n  if (config.smtpCredentialConfigured !== true) issues.push('crea y confirma la credencial SMTP dentro de n8n');\n  if (config.optOutProcessVerified !== true) issues.push('verifica el proceso de bajas y exclusiones por respuesta');\n  if (config.sendLimitsConfirmed !== true) issues.push('confirma el límite operativo diario y mensual');\n  if (config.followUpTimingConfirmed !== true) issues.push('confirma los días entre contactos');\n  if (config.manualBatchControlEnabled !== true) issues.push('habilita la selección manual de leads');\n  if (config.campaignAutoEnabled !== false || config.campaignRotationEnabled !== false || String(config.aiCampaignMode).toLowerCase() !== 'off') issues.push('mantén la asignación de campaña manual y fija por lead');\n  if (issues.length) throw new Error('ENVÍO BLOQUEADO: ' + issues.join('; ') + '.');\n}\nreturn $input.all();`,
    },
    id: '7c835950-f200-4a8c-8e8e-2394b2600001',
    name: preflightName,
    type: 'n8n-nodes-base.code',
    typeVersion: 2,
    position: [540, 220],
  };
  workflow.nodes.push(preflight);
  for (const [, connections] of Object.entries(workflow.connections)) {
    for (const output of connections.main || []) {
      for (const target of output || []) {
        if (target.node === 'Leer Empresas y Archivos') target.node = preflightName;
      }
    }
  }
  workflow.connections[preflightName] = { main: [[{ node: 'Leer Empresas y Archivos', type: 'main', index: 0 }]] };
}

// Lee el marcador operativo únicamente después de que una fila pasó el filtro de selección.
const activeFilter = node('Filtrar Leads Activos');
if (!activeFilter) throw new Error('Active lead filter not found');
const hasBatchFilter = activeFilter.parameters.conditions.conditions.some((c) => String(c.leftValue || '').includes('Enviar ahora'));
if (!hasBatchFilter) activeFilter.parameters.conditions.conditions.push({
  leftValue: '={{ ["si","sí","yes","true","1"].includes(String($json["Enviar ahora"] || "").trim().toLowerCase()) }}',
  rightValue: true,
  operator: { type: 'boolean', operation: 'equals' },
});

// Una selección es de un solo uso: queda consumida al reservar antes del envío SMTP.
const dedup = node('Validar y Deduplicar Leads');
dedup.parameters.jsCode = replaceExactly(
  dedup.parameters.jsCode,
  "accepted.push({ json: row, pairedItem: { item: i } });",
  "accepted.push({ json: { ...row, _batchSelected: true }, pairedItem: { item: i } });",
  'mark selected lead',
);
const reserve = node('Preparar Archivo de Reserva Local');
reserve.parameters.jsCode = replaceExactly(
  reserve.parameters.jsCode,
  "'Requiere revision': 'No',\n  };",
  "'Requiere revision': 'No',\n    'Enviar ahora': row._batchSelected === true ? 'No' : (row['Enviar ahora'] || 'No'),\n  };",
  'consume selection at reservation',
);
const update = node('Preparar Archivo de Actualización Local');
update.parameters.jsCode = replaceExactly(
  update.parameters.jsCode,
  "'Requiere revision': row['Requiere revision'] || 'No',\n    // Campaña no se escribe hasta que exista un ENVÍO confirmado.",
  "'Requiere revision': row['Requiere revision'] || 'No',\n    'Enviar ahora': row._batchSelected === true ? 'No' : (row['Enviar ahora'] || 'No'),\n    // Campaña no se escribe hasta que exista un ENVÍO confirmado.",
  'clear processed batch flag',
);

// Requiere una campaña puesta por la persona que segmentó/seleccionó el registro.
const importPath = path.join(__dirname, '..', 'Importar_Unialiados_Optimizado.ps1');
let importer = fs.readFileSync(importPath, 'utf8');
importer = replaceExactly(
  importer,
  "        if (-not $campaign) { $campaign = 'Unialiados' }\n        if ($validPriorities -notcontains $priority)",
  "        if (-not $campaign) { $campaign = 'Unialiados' }\n        if ($email -and $campaign -eq 'Auto') { $errors.Add(\"Fila $r`: seleccione una campana concreta; la asignacion automatica esta deshabilitada\"); continue }\n        if ($validPriorities -notcontains $priority)",
  'reject automatic campaign assignment for contacts',
);
importer = replaceExactly(
  importer,
  "            'Excluir (Sí/No)' = 'No'\n            'Observaciones comerciales' = $rec.Notes",
  "            'Excluir (Sí/No)' = 'No'\n            'Enviar ahora' = 'No'\n            'Observaciones comerciales' = $rec.Notes",
  'default imported leads to unselected',
);
fs.writeFileSync(importPath, importer, 'utf8');

// Las listas son obligatorias en todas las etapas; el renderer ya las muestra con • en texto
// plano y <ul><li> en HTML. Se exige el mismo contenido por los dos canales.
const prompt = node('Construir Prompt de Correo');
let promptCode = prompt.parameters.jsCode;
promptCode = promptCode.replaceAll('Deja bullets como arreglo vacío [].', 'Incluye exactamente 2 viñetas breves y concretas en bullets; no las repitas en paragraphs ni en cta.');
promptCode = promptCode.replaceAll('Deja bullets como un arreglo vacío [] salvo que la etapa indique explícitamente lo contrario.', 'Incluye exactamente 2 viñetas breves y concretas en bullets; no las escribas dentro de paragraphs.');
promptCode = replaceExactly(
  promptCode,
  'Escribe las viñetas sin guiones, símbolos ni numeración al inicio, y no envuelvas ningún texto entre comillas.',
  'Devuelve exactamente 2 elementos en bullets, cada uno con un beneficio o tarea verificable y distinto; no pongas marcadores al inicio ni repitas frases de paragraphs. El workflow las presenta con el símbolo • en texto plano y como lista HTML. No envuelvas textos entre comillas.',
  'require two concrete bullet points',
);
promptCode = replaceExactly(
  promptCode,
  'bullets (array de 0 a 3 strings)',
  'bullets (array de exactamente 2 strings)',
  'prompt schema description',
);
prompt.parameters.jsCode = promptCode;

const validator = node('Preparar Datos de Envío');
let validatorCode = validator.parameters.jsCode;
validatorCode = replaceExactly(
  validatorCode,
  "else if (!cta) validationError = `Ollama no devolvió CTA válido para ${leadLabel}.`;",
  "else if (!cta) validationError = `Ollama no devolvió CTA válido para ${leadLabel}.`;\n      else if (bullets.length !== 2) validationError = `El correo de ${leadLabel} debe tener exactamente 2 viñetas para que el formato final incluya el símbolo •.`;",
  'validate mandatory bullet count',
);
const duplicateBulletCheck = "      else if (bullets.length !== 2) validationError = `El correo de ${leadLabel} debe tener exactamente 2 viñetas para que el formato final incluya el símbolo •.`;";
validatorCode = validatorCode.replace(`${duplicateBulletCheck}\n${duplicateBulletCheck}`, duplicateBulletCheck);
validatorCode = replaceExactly(
  validatorCode,
  "const ctaWords = plainText(cta).replace(/[^a-z0-9ñ ]+/g, ' ').split(/\\s+/).filter(Boolean);\n  if (ctaWords.length < 4) return false;",
  "const ctaWords = plainText(cta).replace(/[^a-z0-9ñ ]+/g, ' ').split(/\\s+/).filter(Boolean);\n  if (ctaWords.length < 4) return false;",
  'preserve CTA validation anchor',
);
validator.parameters.jsCode = validatorCode;

const generator = node('Generar Correo con IA (Ollama)');
const oldBulletSchema = `"bullets": {\n        "type": "array",\n        "items": { "type": "string" },\n        "minItems": 0,\n        "maxItems": 3\n      }`;
const newBulletSchema = `"bullets": {\n        "type": "array",\n        "items": { "type": "string" },\n        "minItems": 2,\n        "maxItems": 2\n      }`;
generator.parameters.jsonBody = replaceExactly(generator.parameters.jsonBody, oldBulletSchema, newBulletSchema, 'enforce two bullets in Ollama JSON schema');

const campaignValidator = node('Validar Campaña Asignada');
let campaignCode = campaignValidator.parameters.jsCode;
campaignCode = replaceExactly(
  campaignCode,
  `  if (!error && finalKey === 'AXA Colpatria' && config.axaCommercialApproved !== true) {\n    error = \`Lead \${leadLabel}: AXA aún no está autorizada comercialmente. Aprobar con Cristian y activar axaCommercialApproved.\`;\n  }`,
  `  if (!error && finalKey === 'AXA Colpatria' && config.axaCommercialApproved !== true) {\n    error = \`Lead \${leadLabel}: AXA aún no está autorizada comercialmente. Aprobar con Cristian y activar axaCommercialApproved.\`;\n  }\n  if (!error && finalKey === 'AXA Colpatria' && config.axaExclusionsComplete !== true) {\n    error = \`Lead \${leadLabel}: falta confirmar la matriz completa de actividades excluidas de AXA Colpatria.\`;\n  }\n  if (!error && finalKey === 'AXA Colpatria' && !plain(row['Segmento'])) {\n    error = \`Lead \${leadLabel}: AXA requiere un segmento/actividad explícito para revisar elegibilidad.\`;\n  }`,
  'hold AXA until exclusion list and sector are confirmed',
);
campaignCode = campaignCode.replace('Valores permitidos: Unialiados, AXA Colpatria, Unisoft o "Auto".', 'Valores permitidos: Unialiados, AXA Colpatria o Unisoft.');
campaignValidator.parameters.jsCode = campaignCode;

// El modo de rotación automático contradice el acuerdo de escoger y sostener una campaña por lead.
const note = node('NOTA - Arquitectura');
if (note && note.parameters.content && !note.parameters.content.includes('selección manual')) {
  note.parameters.content += '\n\nACUERDO DE REUNIÓN 2026-10-29: el flujo queda con arranque manual y selección explícita por columna «Enviar ahora» en Base de Datos. La campaña se asigna manualmente y se mantiene en la secuencia del lead. El envío está apagado por defecto hasta completar remitente, credencial SMTP, bajas, límites y tiempos de seguimiento.';
}

fs.writeFileSync(file, JSON.stringify(workflow, null, 2) + '\n', 'utf8');
console.log('Updated workflow for manual batches, disabled schedule, production preflight and mandatory bullets.');
