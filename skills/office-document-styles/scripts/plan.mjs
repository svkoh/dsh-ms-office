import process from 'node:process';
import console from 'node:console';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, URL } from 'node:url';
import { randomUUID } from 'node:crypto';

import { buildWordSteps, buildExcelSteps } from './word-excel-recipes.mjs';
import { buildPptSteps } from './ppt-recipes.mjs';

const presetFile = new URL('../resources/presets.json', import.meta.url);
export const presets = JSON.parse(fs.readFileSync(presetFile, 'utf8')).presets;
export function textWidth(text) { return [...text].reduce((n, char) => n + (/^[\x20-\x7e]$/.test(char) ? 0.55 : char === '\n' ? 0 : 1), 0); }
function text(value, label, maximum) {
  if (typeof value !== 'string' || !value.trim() || [...value].some((char) => char.codePointAt(0) < 32)) throw new Error(`${label}: nonempty single-line text required`);
  if (textWidth(value) > maximum) throw new Error(`${label}: content exceeds budget ${maximum}; split paragraphs/slides instead of shrinking or truncating`);
  return value;
}
function object(value, label) { if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`${label}: object required`); }
function noFormulas(value) {
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index++) { if (!Object.hasOwn(value, index)) throw new Error('Sparse arrays are not accepted'); noFormulas(value[index]); }
    return;
  }
  if (!value || typeof value !== 'object') { if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('Non-finite number'); return; }
  for (const [key, item] of Object.entries(value)) {
    if (/formula/i.test(key)) throw new Error('User formulas are not accepted; only recipe-generated safe internal formulas are allowed');
    noFormulas(item);
  }
}
export function escapeCell(value) {
  if (typeof value === 'number') { if (!Number.isFinite(value)) throw new Error('Non-finite cell number'); return value; }
  if (typeof value !== 'string') throw new Error('Cells must be strings or finite numbers');
  return /^[=+\-@]/.test(value) ? `'${value}` : value;
}

/** Results are successful action data objects indexed by step ID, not transport wrappers. */
export function resolveStep(step, results) {
  object(step, 'step'); object(results, 'results');
  const out = structuredClone(step);
  object(out.args, 'step.args');
  if (out.bindings !== undefined && !Array.isArray(out.bindings)) throw new Error('bindings must be an array');
  if (out.bindings?.length && !['ppt_add_elements', 'ppt_inspect_slides', 'ppt_render_slides'].includes(out.action)) throw new Error('Bindings unsupported for action');
  const seen = new Set();
  for (const binding of out.bindings || []) {
    if (!binding || !/^slideId$|^slideIds\.[0-7]$/.test(binding.path) || seen.has(binding.path) || binding.from !== 'add-slides' ||
        !Array.isArray(binding.resultPath) || binding.resultPath.length !== 3 || binding.resultPath[0] !== 'added' ||
        !Number.isInteger(binding.resultPath[1]) || binding.resultPath[1] < 0 || binding.resultPath[1] > 7 || binding.resultPath[2] !== 'id') throw new Error('Unsafe binding');
    if (binding.path.startsWith('slideIds.') && Number(binding.path.split('.')[1]) !== binding.resultPath[1]) throw new Error('Mismatched slide array binding');
    if (!Object.hasOwn(results, binding.from)) throw new Error(`Missing successful result: ${binding.from}`);
    let value = results[binding.from];
    if (value?.ok === false || value?.error) throw new Error('Binding result is not successful');
    if (!Array.isArray(value?.added) || value.added.length !== 8 || value.added.some((slide) => !slide || !Object.hasOwn(slide, 'id') || typeof slide.id !== 'string' || !slide.id.trim()) || new Set(value.added.map((slide) => slide.id)).size !== 8) throw new Error('Expected eight unique successful slide IDs');
    for (const key of binding.resultPath) {
      if (value === null || typeof value !== 'object' || !Object.hasOwn(value, key)) throw new Error('Missing own result binding value');
      value = value[key];
    }
    if (typeof value !== 'string' || !value.trim()) throw new Error('Binding must resolve to a slide ID');
    if (binding.path === 'slideId') {
      if (out.args.slideId !== null) throw new Error('Expected null slideId binding slot');
      out.args.slideId = value;
    } else {
      const index = Number(binding.path.split('.')[1]);
      if (!Array.isArray(out.args.slideIds) || out.args.slideIds[index] !== null) throw new Error('Expected null slideIds binding slot');
      out.args.slideIds[index] = value;
    }
    seen.add(binding.path);
  }
  if (out.action === 'ppt_add_elements' && (typeof out.args.slideId !== 'string' || !out.args.slideId.trim())) throw new Error('Unresolved slide binding');
  if (['ppt_inspect_slides', 'ppt_render_slides'].includes(out.action) && (!Array.isArray(out.args.slideIds) || out.args.slideIds.length !== 8 || out.args.slideIds.some((id) => typeof id !== 'string' || !id.trim()) || new Set(out.args.slideIds).size !== 8)) throw new Error('Unresolved or duplicate slide bindings');
  if (out.action === 'ppt_add_elements' && Array.isArray(out.args.elements)) for (const element of out.args.elements) delete element.role;
  delete out.bindings;
  return out;
}

export function buildPlan(presetId, input, instance) {
  const found = presets.find((item) => item.id === presetId);
  if (!found) throw new Error(`Unknown preset: ${presetId}`);
  const preset = structuredClone(found);
  object(input, 'input'); noFormulas(input);
  if (typeof instance !== 'string' || !instance.trim() || instance.length > 200) throw new Error('Explicit Office instance is required');
  if (!['USER', 'DEMO / 示例数据'].includes(input.dataLabel)) throw new Error('dataLabel must be USER or DEMO / 示例数据');
  let sourceLabel = '';
  if (input.sourceLabel !== undefined && input.sourceLabel !== '') sourceLabel = text(input.sourceLabel, 'sourceLabel', 30);
  const displayLabel = input.dataLabel === 'DEMO / 示例数据' ? 'DEMO / 示例数据' : sourceLabel;
  const plan = { schemaVersion: 1, planId: randomUUID(), host: preset.host, presetId, instance, dataLabel: input.dataLabel, displayLabel, preset,
    preconditions: ['Confirm the intended Office instance and user authorization before executing.', 'Check installed fonts and host API requirements; never download fonts.', 'Execute steps sequentially; inject top-level instance and derive idempotencyKey from planId + step.id.'],
    manualChecks: ['Plan checks do not verify real Office rendering. Read back content and visually inspect the actual document.'], steps: [] };
  const add = (id, action, args, bindings) => plan.steps.push({ id, action, args, idempotencyKey: `${plan.planId}:${id}`, ...(bindings ? { bindings } : {}) });
  if (preset.host === 'Word') buildWordSteps(preset,input,add,plan);
  else if (preset.host === 'Excel') buildExcelSteps(preset,input,add,plan);
  else buildPptSteps(preset,input,add,plan);
  return plan;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const [, , command, id, inputFile, instance] = process.argv;
    if (command === 'list') console.log(JSON.stringify({ schemaVersion: 1, presets: presets.map(({ id, host, name }) => ({ id, host, name })) }, null, 2));
    else if (command === 'build' && id && inputFile && instance) console.log(JSON.stringify(buildPlan(id, JSON.parse(fs.readFileSync(inputFile, 'utf8')), instance), null, 2));
    else throw new Error('Usage: node plan.mjs list | build <preset-id> <input.json> <instance>');
  } catch (error) { console.error(error.message); process.exitCode = 1; }
}
