'use strict';
// Format-independent entry points used by the app (and the tests).
const path = require('path');
const docx = require('./docx');
const pptx = require('./pptx');
const pdf = require('./pdf');
const { humanChecks } = require('./issues');

const HANDLERS = { docx, pptx, pdf };
const SUPPORTED = ['.docx', '.pptx', '.pdf'];

function detectFormat(fileName) {
  const ext = path.extname(fileName || '').toLowerCase();
  if (ext === '.docx') return 'docx';
  if (ext === '.pptx') return 'pptx';
  if (ext === '.pdf') return 'pdf';
  if (ext === '.doc' || ext === '.ppt') throw new Error(`${ext} is an old binary format. Save it as ${ext}x in Office first, then add it here.`);
  throw new Error(`Unsupported file type ${ext || '(none)'}. Use .docx, .pptx or .pdf.`);
}

function summarise(issues) {
  const s = { errors: 0, warnings: 0, auto: 0, ai: 0, input: 0, manual: 0, total: issues.length };
  for (const i of issues) {
    if (i.severity === 'error') s.errors++; else s.warnings++;
    s[i.fix.kind] = (s[i.fix.kind] || 0) + 1;
  }
  s.passesAutomated = s.errors === 0;
  return s;
}

async function audit(buffer, fileName, opts = {}) {
  const format = detectFormat(fileName);
  const result = await HANDLERS[format].audit(buffer, { ...opts, fileName });
  result.fileName = fileName;
  result.summary = summarise(result.issues);
  result.humanChecks = humanChecks(format);
  result.checkedAt = new Date().toISOString();
  return result;
}

async function applyFix(buffer, fileName, issue, input = {}, opts = {}) {
  const format = detectFormat(fileName);
  return HANDLERS[format].fix(buffer, issue, input, opts);
}

/** Issues that "Fix all" can apply with suggested text (offline or AI) when asked to. */
function suggestible(issue, extra) {
  if (!['ai', 'input'].includes(issue.fix.kind) || issue.fix.inFixAll === false) return false;
  const x = extra && extra[issue.id];
  if (x && (x.decorative || (x.text || '').trim())) return true;
  return !!(issue.fix.suggestion || '').trim() && issue.fix.inputType !== 'level';
}

/**
 * Apply every automatic fix, re-checking after each so locations stay correct.
 * With { includeSuggestions: true } it also applies suggested text (captions, chart data,
 * link addresses, first slide line, and any AI suggestions passed in `suggestions` by issue id).
 * Those are recorded with `suggested: true` so a person can review them.
 */
async function fixAllAuto(buffer, fileName, opts = {}, onProgress, mode = {}) {
  let buf = buffer;
  const applied = [], failed = new Set(), errors = [];
  const extra = mode.suggestions || {};
  const pick = (res) => res.issues.find((i) => i.fix.kind === 'auto' && i.fix.inFixAll !== false && !failed.has(i.id))
    || (mode.includeSuggestions ? res.issues.find((i) => !failed.has(i.id) && suggestible(i, extra)) : null);
  for (let guard = 0; guard < 300; guard++) {
    const res = await audit(buf, fileName, opts);
    const next = pick(res);
    if (!next) return { buffer: buf, applied, errors, result: res };
    const isSuggestion = next.fix.kind !== 'auto';
    const x = extra[next.id];
    const input = !isSuggestion ? {} : x && x.decorative ? { decorative: true } : { value: ((x && x.text) || next.fix.suggestion).trim() };
    try {
      buf = await applyFix(buf, fileName, next, input, opts);
      if (isSuggestion) failed.add(next.id); // apply suggested text to an item once per run
      const item = { id: next.id, title: next.title };
      if (isSuggestion) {
        item.suggested = true;
        item.source = x ? x.source || 'AI suggestion' : next.fix.suggestionSource || 'suggested text';
        item.detail = input.decorative ? 'Marked as decorative' : `“${input.value}”`;
        item.location = next.location && next.location.label;
      }
      applied.push(item);
      if (onProgress) onProgress(applied.length, next);
    } catch (e) {
      failed.add(next.id);
      errors.push({ id: next.id, title: next.title, message: e.message });
    }
  }
  const res = await audit(buf, fileName, opts);
  return { buffer: buf, applied, errors, result: res };
}

async function getImage(buffer, fileName, target) {
  const format = detectFormat(fileName);
  return HANDLERS[format].getImage ? HANDLERS[format].getImage(buffer, target) : null;
}

/** "Report (accessible).docx" style name for the remediated copy. */
function fixedName(fileName) {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext).replace(/ \(accessible\)$/, '');
  return `${base} (accessible)${ext}`;
}

module.exports = { SUPPORTED, detectFormat, audit, applyFix, fixAllAuto, suggestible, getImage, fixedName, summarise };
