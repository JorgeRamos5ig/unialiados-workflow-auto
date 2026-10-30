const fs = require('fs');
const workflow = JSON.parse(fs.readFileSync(require('path').join(__dirname, '..', 'Unialiados - Prospección local (v17.4 fix gramatical).json'), 'utf8'));
const names = new Set(workflow.nodes.map((node) => node.name));
let codeNodes = 0;
for (const node of workflow.nodes) {
  if (node.parameters?.jsCode) {
    new Function('$input', '$', '$execution', '$getWorkflowStaticData', '$json', node.parameters.jsCode);
    codeNodes += 1;
  }
}
const missingEdges = [];
for (const [source, groups] of Object.entries(workflow.connections || {})) {
  if (!names.has(source)) missingEdges.push(`source:${source}`);
  for (const group of ['main', 'ai', 'error']) {
    for (const branch of groups[group] || []) {
      for (const edge of branch) if (!names.has(edge.node)) missingEdges.push(`${source}->${edge.node}`);
    }
  }
}
const configNode = workflow.nodes.find((node) => node.name === 'Config: batch y remitente');
const sendEnabled = configNode?.parameters.assignments.assignments.find((assignment) => assignment.name === 'sendEnabled')?.value;
const schedule = workflow.nodes.find((node) => node.name.startsWith('Programado'));
const result = { nodes: workflow.nodes.length, codeNodes, missingEdges, active: workflow.active === true, sendEnabled, scheduleDisabled: schedule?.disabled === true };
console.log(JSON.stringify(result));
if (missingEdges.length || result.active || sendEnabled !== false || !result.scheduleDisabled) process.exit(1);
