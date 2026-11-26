// Unit tests for the audit/fix engine. Run: npm test
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const engine = require('../src/core/engine');
const C = require('../src/core/contrast');
const { buildReport } = require('../src/core/report');

const fx = (n) => fs.readFileSync(path.join(__dirname, 'fixtures', n));
const rules = (r) => r.issues.map((i) => i.rule);

test('contrast maths matches WCAG reference values', () => {
  assert.ok(Math.abs(C.ratio('000000', 'FFFFFF') - 21) < 0.01);
  assert.ok(Math.abs(C.ratio('777777', 'FFFFFF') - 4.48) < 0.01);
  const fixed = C.fixForeground('AAAAAA', 'FFFFFF', 4.5);
  assert.ok(C.ratio(fixed, 'FFFFFF') >= 4.5);
  assert.strictEqual(C.requiredRatio(18, false), 3);
  assert.strictEqual(C.requiredRatio(14, true), 3);
  assert.strictEqual(C.requiredRatio(12, true), 4.5);
});

test('Word: finds the expected problems', async () => {
  const r = await engine.audit(fx('sample-inaccessible.docx'), 'sample-inaccessible.docx');
  for (const rule of ['doc-title', 'doc-lang', 'img-alt', 'heading-skip', 'heading-fake', 'table-header', 'contrast', 'list-manual', 'link-text', 'blank-paras']) {
    assert.ok(rules(r).includes(rule), `expected ${rule}`);
  }
  for (const i of r.issues) assert.ok(i.wcag.length && i.wcag[0].name, `${i.rule} has a WCAG mapping`);
});

test('Word: fix all automatic + text fixes passes automated checks and stays a valid file', async () => {
  const name = 'sample-inaccessible.docx';
  let { buffer, result } = await engine.fixAllAuto(fx(name), name, { defaultLanguage: 'en-GB' });
  assert.ok(!result.issues.some((i) => i.fix.kind === 'auto'), 'no automatic issues left');
  const img = result.issues.find((i) => i.rule === 'img-alt');
  buffer = await engine.applyFix(buffer, name, img, { value: 'Bar chart of students per region' });
  let r = await engine.audit(buffer, name);
  const link = r.issues.find((i) => i.rule === 'link-text');
  buffer = await engine.applyFix(buffer, name, link, { value: 'English programme data 2026' });
  r = await engine.audit(buffer, name);
  assert.strictEqual(r.summary.errors, 0, JSON.stringify(r.issues.map((i) => i.title)));
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /descr="Bar chart of students per region"/);
  assert.match(xml, /<w:tblHeader\/>/);
  assert.match(xml, /<w:numPr>/);
  assert.ok(zip.file('word/numbering.xml'), 'numbering part created');
  const core = await zip.file('docProps/core.xml').async('string');
  assert.match(core, /Annual English Programme Report/);
});

test('Word: marking an image decorative clears the issue', async () => {
  const name = 'sample-inaccessible.docx';
  const r = await engine.audit(fx(name), name);
  const img = r.issues.find((i) => i.rule === 'img-alt');
  const buf = await engine.applyFix(fx(name), name, img, { decorative: true });
  const r2 = await engine.audit(buf, name);
  assert.ok(!rules(r2).includes('img-alt'));
});

test('PowerPoint: finds and fixes problems', async () => {
  const name = 'sample-inaccessible.pptx';
  const r = await engine.audit(fx(name), name);
  for (const rule of ['doc-title', 'doc-lang', 'slide-title', 'img-alt-quality', 'contrast', 'slide-title-dup', 'table-header', 'link-text']) assert.ok(rules(r).includes(rule), `expected ${rule}`);
  let { buffer, result } = await engine.fixAllAuto(fx(name), name, {});
  const st = result.issues.find((i) => i.rule === 'slide-title');
  buffer = await engine.applyFix(buffer, name, st, { value: 'Our results this year' });
  const r2 = await engine.audit(buffer, name);
  assert.ok(!rules(r2).includes('slide-title'));
  const zip = await JSZip.loadAsync(buffer);
  const s2 = await zip.file('ppt/slides/slide2.xml').async('string');
  assert.match(s2, /type="title"/);
  assert.match(s2, /lang="en-GB"/);
});

test('PDF: untagged file is flagged honestly and metadata fixes apply', async () => {
  const name = 'sample-untagged.pdf';
  const r = await engine.audit(fx(name), name);
  const untagged = r.issues.find((i) => i.rule === 'untagged');
  assert.ok(untagged && untagged.fix.kind === 'manual', 'untagged PDFs are not auto-tagged');
  const { result } = await engine.fixAllAuto(fx(name), name, { defaultLanguage: 'en-GB' });
  assert.deepStrictEqual(rules(result).sort(), ['fonts', 'untagged']);
  assert.strictEqual(result.meta.language, 'en-GB');
});

test('Originals are never modified by fixes', async () => {
  const name = 'sample-inaccessible.docx';
  const orig = fx(name);
  const copy = Buffer.from(orig);
  await engine.fixAllAuto(orig, name, {});
  assert.strictEqual(Buffer.compare(orig, copy), 0);
  assert.strictEqual(engine.fixedName('Report.docx'), 'Report (accessible).docx');
});

test('Report lists remaining issues and human checks', async () => {
  const name = 'sample-inaccessible.docx';
  const r = await engine.audit(fx(name), name);
  const html = buildReport({ fileName: name, format: 'docx', original: r, current: r, applied: [], humanChecks: {} });
  assert.match(html, /<html lang="en-GB">/);
  assert.match(html, /Does not pass/);
  assert.match(html, /Reading order makes sense/);
});

test('Unsupported formats give a clear message', () => {
  assert.throws(() => engine.detectFormat('old.doc'), /Save it as .docx/);
});
