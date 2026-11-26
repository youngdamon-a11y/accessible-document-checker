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

/** Apply every automatic fix, re-checking after each so locations stay correct. */
async function fixAllAuto(buffer, fileName, opts = {}, onProgress) {
  let buf = buffer;
  const applied = [], failed = new Set(), errors = [];
  for (let guard = 0; guard < 200; guard++) {
    const res = await audit(buf, fileName, opts);
    const next = res.issues.find((i) => i.fix.kind === 'auto' && i.fix.inFixAll !== false && !failed.has(i.id));
    if (!next) return { buffer: buf, applied, errors, result: res };
    try {
      buf = await applyFix(buf, fileName, next, {}, opts);
      applied.push({ id: next.id, title: next.title });
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

module.exports = { SUPPORTED, detectFormat, audit, applyFix, fixAllAuto, getImage, fixedName, summarise };
