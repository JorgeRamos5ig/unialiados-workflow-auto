const fs = require('fs');

const WORKFLOW = require('path').join(__dirname, '..', 'Unialiados - Prospección local (v17.4 fix gramatical).json');
const API = 'http://127.0.0.1:11434/api/chat';
const MODEL = 'qwen3.5:4b';
const data = JSON.parse(fs.readFileSync(WORKFLOW, 'utf8'));
const node = (name) => data.nodes.find((n) => n.name === name);
const configNode = node('Config: batch y remitente');
const promptNode = node('Construir Prompt de Correo');
const validatorNode = node('Preparar Datos de Envío');

function configFromNode() {
  const out = {};
  for (const a of configNode.parameters.assignments.assignments) {
    if (typeof a.value !== 'string' || !a.value.startsWith('=')) out[a.name] = a.value;
  }
  out.reportDate = new Date().toISOString();
  out.runBatchId = 'acceptance-test';
  return out;
}

const CONFIG = configFromNode();
const state = {};
function makeDollar(rows) {
  return (name) => {
    const list = rows[name] || [];
    return {
      first: () => list[0],
      all: () => list,
      itemMatching: (i) => list[i],
    };
  };
}
function runCode(code, input, rows) {
  const $ = makeDollar(rows);
  const $input = { all: () => input.map((json) => ({ json })) };
  const execution = { id: 'acceptance-test' };
  const getStatic = () => state;
  const fn = new Function('$input', '$', '$execution', '$getWorkflowStaticData', '$json', code);
  return fn($input, $, execution, getStatic, input[0]);
}

const base = {
  ID_Lead: 9000,
  ID_Empresa: 9000,
  Empresa: 'Empresa Demo Colombiana',
  Contacto: 'María Rodríguez',
  Cargo: 'Directora administrativa',
  Correo: 'prueba@example.invalid',
  Segmento: 'empresa de servicios y contratistas',
  Estrategia: 'Validar cómo organiza hoy sus procesos administrativos.',
  'Próximo paso': 'Compartir información y conversar cuando sea oportuno.',
  'Observaciones comerciales': '',
  'Campaña asignada': 'Auto',
  'Historial de correos': '',
  _campaignKey: '',
  _campaignValid: true,
  _campaignAuto: true,
  _campaignAxaEligibleToDiscuss: true,
  _campaignHold: '',
  _campaignAIAllowed: [],
  _campaignAIDefault: '',
  _aiAttempt: 1,
};

const cases = [
  ['Unialiados', 'FIRST', 'administración de Seguridad Social, novedades y cumplimiento laboral'],
  ['Unialiados', 'SECOND', 'seguimiento de nómina y novedades laborales'],
  ['Unialiados', 'THIRD', 'organización del proceso de nómina'],
  ['AXA Colpatria', 'FIRST', 'revisión de la ARL actual, prevención y riesgos laborales'],
  ['AXA Colpatria', 'SECOND', 'acompañamiento preventivo y situación actual de ARL'],
  ['AXA Colpatria', 'THIRD', 'evaluar una conversación sobre ARL'],
  ['Unisoft', 'FIRST', 'documentos laborales dispersos en varias sedes y control de vigencias'],
  ['Unisoft', 'SECOND', 'consulta y trazabilidad documental de trabajadores'],
  ['Unisoft', 'THIRD', 'organización de archivos laborales'],
];

function words(text) { return String(text || '').trim().split(/\s+/).filter(Boolean).length; }

async function askOllama(system, user) {
  const response = await fetch(API, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      model: MODEL,
      messages: [{ role: 'system', content: system }, { role: 'user', content: user }],
      stream: false,
      think: false,
      format: {
        type: 'object',
        properties: {
          subject: { type: 'string' },
          paragraphs: { type: 'array', items: { type: 'string' }, minItems: 1, maxItems: 3 },
          bullets: { type: 'array', items: { type: 'string' }, minItems: 0, maxItems: 3 },
          cta: { type: 'string' },
        },
        required: ['subject', 'paragraphs', 'bullets', 'cta'],
        additionalProperties: false,
      },
      options: { temperature: 0.35, top_p: 0.85, repeat_penalty: 1.08, num_ctx: 8192, num_predict: 700 },
    }),
  });
  if (!response.ok) throw new Error(`Ollama HTTP ${response.status}: ${await response.text()}`);
  return response.json();
}

async function main() {
  const results = [];
  const selectedCases = process.argv[2]
    ? cases.filter(([campaign, stage]) => `${campaign}/${stage}` === process.argv[2])
    : cases;
  for (const [campaign, stage, signal] of selectedCases) {
    let row = { ...base, ID_Lead: 9000 + results.length, 'Campaña asignada': campaign, _campaignKey: campaign, _stage: stage };
    row.Estrategia = signal;
    let prepared;
    let checked;
    let attempts = 0;
    let rawContent = '';
    while (attempts < Number(CONFIG.maxAiAttempts || 3)) {
      attempts += 1;
      const promptRows = {
        'Config: batch y remitente': [{ json: CONFIG }],
        'Extraer Empresas y Archivos': [{ json: { Empresa: row.Empresa, 'Campaña principal': campaign, Sector: row.Segmento } }],
      };
      prepared = runCode(promptNode.parameters.jsCode, [row], promptRows)[0].json;
      const raw = await askOllama(prepared._systemPrompt, prepared._userPrompt);
      rawContent = raw.message?.content || '';
      const validatorRows = {
        'Config: batch y remitente': [{ json: CONFIG }],
        'Construir Prompt de Correo': [{ json: prepared }],
      };
      checked = runCode(validatorNode.parameters.jsCode, [{ message: { content: rawContent } }], validatorRows)[0].json;
      if (checked._validationOk === true) break;
      row = { ...prepared, _validationError: checked._validationError, _aiAttempt: attempts + 1 };
    }
    const body = [checked._emailSubject, checked._emailBody, checked._emailBodyHtml].join('\n');
    const result = {
      campaign,
      stage,
      ok: checked._validationOk === true,
      attempts,
      error: checked._validationError || '',
      words: words(checked._emailBody),
      subject: checked._emailSubject || '',
      body: checked._emailBody || '',
      raw: checked._validationOk === true ? '' : rawContent,
      advertisingSignals: /(oferta imperdible|promoción|descuento|aproveche|última oportunidad|compr[aá] ahora|garantizamos|somos los mejores|líderes? del mercado|cotización personalizada)/i.test(body),
    };
    results.push(result);
    console.log(JSON.stringify(result));
  }
  const passed = results.filter((r) => r.ok).length;
  const clean = results.filter((r) => !r.advertisingSignals).length;
  console.log(JSON.stringify({ summary: { total: results.length, passed, rejected: results.length - passed, nonAdvertising: clean } }));
  if (passed !== results.length || clean !== results.length) process.exitCode = 1;
}

main().catch((error) => { console.error(error.stack || error); process.exitCode = 2; });
