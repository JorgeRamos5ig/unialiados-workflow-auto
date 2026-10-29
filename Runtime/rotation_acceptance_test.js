const fs = require('fs');
const path = require('path');
const workflow = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'Unialiados - Prospección local (v17.4 fix gramatical).json'), 'utf8'));
const getNode = (name) => workflow.nodes.find((n) => n.name === name);
const config = {};
for (const a of getNode('Config: batch y remitente').parameters.assignments.assignments) {
  if (typeof a.value !== 'string' || !a.value.startsWith('=')) config[a.name] = a.value;
}
config.campaignRotationEnabled = true;
config.aiCampaignMode = 'always';
config.axaAutoAssignEnabled = true;
config.axaCommercialApproved = true;

const state = {};
function dollar(rows) {
  return (name) => {
    const list = rows[name] || [];
    return { first: () => list[0], all: () => list, itemMatching: (i) => list[i] };
  };
}
function run(code, input, rows) {
  const $input = { all: () => input.map((json) => ({ json })) };
  return new Function('$input', '$', '$execution', '$getWorkflowStaticData', '$json', code)(
    $input, dollar(rows), { id: 'rotation-acceptance' }, () => state, input[0]
  );
}

const assignCode = getNode('Asignar Campaña (Reglas)').parameters.jsCode;
const applyCode = getNode('Aplicar Decisión de Campaña').parameters.jsCode;
const validateCode = getNode('Validar Campaña Asignada').parameters.jsCode;
const company = { ID_Empresa: 700, Empresa: 'Empresa Rotación Demo', Sector: 'servicios', 'Actividad económica': 'servicios empresariales' };
const base = {
  ID_Lead: 700, ID_Empresa: 700, Empresa: company.Empresa, Contacto: 'María Rodríguez',
  Cargo: 'Directora administrativa', Correo: 'rotation@example.invalid', Segmento: 'servicios empresariales',
  'Campaña asignada': 'Auto', Estrategia: 'revisar procesos de gestión administrativa',
  'Próximo paso': '', 'Observaciones comerciales': '', 'Historial de correos': '',
};

const choices = [];
let row = { ...base };
for (const stage of ['FIRST', 'SECOND', 'THIRD']) {
  const assigned = run(assignCode, [row], {
    'Config: batch y remitente': [{ json: config }],
    'Extraer Base de Leads': [{ json: row }],
    'Extraer Empresas y Archivos': [{ json: company }],
  })[0].json;
  const allowed = assigned._campaignAIAllowed;
  if (!Array.isArray(allowed) || allowed.length !== 3 - choices.length) throw new Error(`unexpected allowed campaigns at ${stage}: ${allowed}`);
  const chosen = allowed[0];
  const applied = run(applyCode, [{ message: { content: JSON.stringify({ campaign: chosen, reason: 'acceptance test' }) } }], {
    '¿Campaña Necesita IA?': [{ json: assigned }],
  })[0].json;
  const validated = run(validateCode, [applied], { 'Config: batch y remitente': [{ json: config }] })[0].json;
  if (!validated._campaignValid) throw new Error(`${stage} rejected: ${validated._validationError}`);
  choices.push(chosen);
  row = {
    ...row,
    'Campaña asignada': 'Auto',
    'Historial de correos': `${row['Historial de correos']}[2026-10-29T12:00:00.000Z] ${stage} CONTACTO\nCampaña: ${chosen}\n`,
  };
}

if (new Set(choices).size !== 3) throw new Error(`rotation repeated a campaign: ${choices.join(', ')}`);
console.log(JSON.stringify({ passed: true, stages: ['FIRST', 'SECOND', 'THIRD'], campaigns: choices }));
