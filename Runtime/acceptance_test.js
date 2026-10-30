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
  const wrappedInput = input.map((json) => ({ json }));
  const $input = { all: () => wrappedInput, first: () => wrappedInput[0] };
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

function runReadinessChecks() {
  const findNode = (name) => data.nodes.find((entry) => entry.name === name);
  const manualTrigger = findNode('Iniciar manualmente');
  const scheduleTrigger = data.nodes.find((entry) => entry.name.startsWith('Programado'));
  const preflight = data.nodes.find((entry) => entry.name.startsWith('Validar configuración'));
  const filter = findNode('Filtrar Leads Activos');
  if (!manualTrigger || !scheduleTrigger?.disabled || !preflight) throw new Error('Manual-only trigger/preflight setup is incomplete.');
  if (!filter.parameters.conditions.conditions.some((condition) => String(condition.leftValue).includes('Enviar ahora'))) {
    throw new Error('Active lead filter does not require the manual batch checkbox.');
  }
  if (CONFIG.sendEnabled !== false || CONFIG.senderEmail !== 'PENDIENTE_CONFIGURAR') {
    throw new Error('Outbound sending must remain disabled until the customer configures it.');
  }
  const blocked = runCode(preflight.parameters.jsCode, [CONFIG], {});
  if (!blocked?.length) throw new Error('Preflight did not return its disabled-mode item.');
  const readyConfig = {
    ...CONFIG,
    sendEnabled: true,
    senderEmail: 'pruebas@unialiados.com',
    senderIdentityConfirmed: true,
    smtpCredentialConfigured: true,
    optOutProcessVerified: true,
    sendLimitsConfirmed: true,
    followUpTimingConfirmed: true,
    manualBatchControlEnabled: true,
    campaignAutoEnabled: false,
    campaignRotationEnabled: false,
    aiCampaignMode: 'off',
  };
  runCode(preflight.parameters.jsCode, [readyConfig], {});
}

function runValidatorRegressionChecks() {
  const validatorCode = validatorNode.parameters.jsCode;
  const testLead = { ...base, _campaignKey: 'Unialiados', _stage: 'FIRST', _campaignValid: true };
  const responseFor = (body) => ({ message: { content: JSON.stringify(body) } });
  const validatorRows = {
    'Config: batch y remitente': [{ json: CONFIG }],
    'Construir Prompt de Correo': [{ json: testLead }],
  };
  const spanishWithWords = {
    subject: 'Organización de novedades laborales',
    paragraphs: [
      'Esta coordinación también puede facilitar un registro ordenado de novedades laborales cuando la operación reúne personal en distintas sedes. El equipo puede centralizar altas, retiros, incapacidades y soportes de aportes en un flujo consultable, manteniendo cada documento asociado a la persona y al periodo correspondiente.',
      'Además, la gestión administrativa puede incluir la revisión de afiliaciones, la preparación de la planilla PILA y la organización de comprobantes para cada trabajador. Esta forma de trabajo también permite que las áreas involucradas consulten el historial sin buscar la información en varios correos o carpetas. El registro conserva cada soporte para una revisión posterior.',
    ],
    bullets: ['Novedades y soportes organizados por trabajador y periodo', 'Afiliaciones y planillas disponibles para consulta del equipo'],
    cta: '¿Le sería útil revisar cómo funciona este acompañamiento en su operación?',
  };
  const output = runCode(validatorCode, [responseFor(spanishWithWords)], validatorRows)[0].json;
  if (output._validationOk !== true) throw new Error(`Spanish-word regression failed: ${output._validationError}`);

  const englishLeak = { ...spanishWithWords, paragraphs: ['In el sector industrial, el equipo organiza sus novedades y soportes. ' + spanishWithWords.paragraphs[1]] };
  const englishOutput = runCode(validatorCode, [responseFor(englishLeak)], validatorRows)[0].json;
  if (englishOutput._validationOk === true) throw new Error(`English fragment regression was not detected: ${englishOutput._validationError}`);

  const duplicatedCta = { ...spanishWithWords, paragraphs: [spanishWithWords.paragraphs[0], spanishWithWords.cta] };
  const ctaOutput = runCode(validatorCode, [responseFor(duplicatedCta)], validatorRows)[0].json;
  if (!/repite la pregunta del CTA/i.test(ctaOutput._validationError || '')) throw new Error('Duplicated CTA regression was not detected.');

  const inventedPain = { ...spanishWithWords, paragraphs: ['Sabía que en su equipo se pierde tiempo organizando los soportes laborales. ' + spanishWithWords.paragraphs[0], spanishWithWords.paragraphs[1]] };
  const inventedRows = { ...validatorRows, 'Construir Prompt de Correo': [{ json: { ...testLead, _campaignKey: 'Unialiados' } }] };
  const inventedOutput = runCode(validatorCode, [responseFor(inventedPain)], inventedRows)[0].json;
  if (!/afirmaci[oó]n no sustentada/i.test(inventedOutput._validationError || '')) throw new Error('Unsupported prospect assumption regression was not detected.');

  const inferredPain = { ...spanishWithWords, paragraphs: ['Escribí porque noto que en empresas como la suya se pierde tiempo buscando archivos. ' + spanishWithWords.paragraphs[0], spanishWithWords.paragraphs[1]] };
  const inferredOutput = runCode(validatorCode, [responseFor(inferredPain)], inventedRows)[0].json;
  if (!/afirmaci[oó]n no sustentada/i.test(inferredOutput._validationError || '')) throw new Error('Implied prospect assumption regression was not detected.');

  const knowledgeClaim = { ...spanishWithWords, paragraphs: ['María, noto que el equipo tiene archivos dispersos y pierde tiempo en cada consulta. ' + spanishWithWords.paragraphs[0], spanishWithWords.paragraphs[1]] };
  const knowledgeOutput = runCode(validatorCode, [responseFor(knowledgeClaim)], inventedRows)[0].json;
  if (!/afirmaci[oó]n no sustentada/i.test(knowledgeOutput._validationError || '')) throw new Error('Unverified knowledge-claim regression was not detected.');

  const axaClaim = { ...spanishWithWords, bullets: ['Disponibilidad del respaldo de un grupo internacional AXA Colpatria', spanishWithWords.bullets[1]] };
  const axaRows = { ...validatorRows, 'Construir Prompt de Correo': [{ json: { ...testLead, _campaignKey: 'AXA Colpatria' } }] };
  const axaOutput = runCode(validatorCode, [responseFor(axaClaim)], axaRows)[0].json;
  if (!/atributo comercial no confirmado/i.test(axaOutput._validationError || '')) throw new Error('Unconfirmed AXA claim regression was not detected.');

  const crossCampaign = { ...spanishWithWords, paragraphs: ['En Unialiados utilizamos UNISOFT para centralizar sus documentos laborales. ' + spanishWithWords.paragraphs[0], spanishWithWords.paragraphs[1]] };
  const crossCampaignOutput = runCode(validatorCode, [responseFor(crossCampaign)], validatorRows)[0].json;
  if (!/mezcla el contenido de otra campaña/i.test(crossCampaignOutput._validationError || '')) throw new Error('Cross-campaign content regression was not detected.');

  const noBullets = { ...spanishWithWords, bullets: [] };
  const noBulletOutput = runCode(validatorCode, [responseFor(noBullets)], validatorRows)[0].json;
  if (!/exactamente 2 viñetas/i.test(noBulletOutput._validationError || '')) throw new Error('Missing bullets were not rejected.');
}

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
          bullets: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 2 },
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
  runReadinessChecks();
  runValidatorRegressionChecks();
  console.log(JSON.stringify({ localReadinessAndValidatorRegressions: 'PASS' }));
  const results = [];
  const requestedCases = process.argv[2] ? new Set(process.argv[2].split(',').map((value) => value.trim())) : null;
  const selectedCases = requestedCases
    ? cases.filter(([campaign, stage]) => requestedCases.has(`${campaign}/${stage}`))
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
    const bullet = String.fromCodePoint(0x2022);
    const bulletLines = String(checked._emailBody || '').split(/\r?\n/).filter((line) => line.startsWith(`${bullet} `));
    const htmlListItems = (String(checked._emailBodyHtml || '').match(/<li\b/gi) || []).length;
    const result = {
      campaign,
      stage,
      ok: checked._validationOk === true && bulletLines.length === 2 && htmlListItems === 2,
      attempts,
      error: checked._validationError || '',
      words: words(checked._emailBody),
      bulletLines: bulletLines.length,
      htmlListItems,
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
  console.log(JSON.stringify({ summary: { total: results.length, passed, rejected: results.length - passed, nonAdvertising: clean, twoPlainTextAndHtmlBullets: results.filter((r) => r.bulletLines === 2 && r.htmlListItems === 2).length } }));
  if (passed !== results.length || clean !== results.length) process.exitCode = 1;
}

main().catch((error) => { console.error(error.stack || error); process.exitCode = 2; });
