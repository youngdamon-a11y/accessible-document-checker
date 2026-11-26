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

// ---------------------------------------------------------------------------
// Offline suggestions and "Fix all, including suggested text"
const D = require('../src/core/describe');

test('Chart alt text is written from the chart data', () => {
  const xml = `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Visitors by month</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Visitors</c:v></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>100</c:v></c:pt><c:pt idx="1"><c:v>250</c:v></c:pt><c:pt idx="2"><c:v>400</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;
  const d = D.describeChart(xml);
  assert.match(d, /^Line chart: Visitors by month\./);
  assert.match(d, /Jan 100, Feb 250, Mar 400/);
  assert.match(d, /rises from 100 to 400/);
  assert.equal(D.describeChart('<not-a-chart/>'), '');
});

test('Captions are recognised and the label removed', () => {
  assert.equal(D.captionText('Figure 3: Map of our offices'), 'Map of our offices');
  assert.equal(D.captionText('Photo – Students in class'), 'Students in class');
  assert.equal(D.captionText('Figures show growth'), '');
  assert.equal(D.findCaption(['', 'Normal text', 'Chart 1. Sales by year']), 'Sales by year');
});

test('Fix all with suggested text fills in captions, chart data, links and titles, and marks them for review', async () => {
  for (const f of ['sample-inaccessible.docx', 'sample-inaccessible.pptx']) {
    const buf = fx(f);
    const plain = await engine.fixAllAuto(buf, f);
    assert.ok(plain.applied.every((a) => !a.suggested), 'plain Fix all must not apply suggested text');
    const r = await engine.fixAllAuto(buf, f, {}, null, { includeSuggestions: true, suggestions: {} });
    const sugg = r.applied.filter((a) => a.suggested);
    assert.ok(sugg.length >= 2, `${f}: expected suggested fixes`);
    assert.ok(sugg.every((a) => a.source && a.detail), 'each suggested fix records its source and text');
    assert.ok(!r.result.issues.some((i) => i.rule === 'link-text'), `${f}: link text fixed`);
  }
  // pptx: the native chart gets alt text from its data
  const pp = await engine.fixAllAuto(fx('sample-inaccessible.pptx'), 'sample-inaccessible.pptx', {}, null, { includeSuggestions: true });
  assert.ok(pp.applied.some((a) => a.source === 'chart data' && /Column chart: Pass rate by region/.test(a.detail)));
});

test('AI suggestions passed to Fix all are used for pictures without captions', async () => {
  const f = 'sample-inaccessible.pptx';
  const buf = fx(f);
  const before = await engine.audit(buf, f);
  const pic = before.issues.find((i) => /^img-alt/.test(i.rule) && i.fix.kind === 'ai' && !i.fix.suggestion);
  assert.ok(pic, 'fixture has a picture with no offline suggestion');
  const r = await engine.fixAllAuto(buf, f, {}, null, { includeSuggestions: true, suggestions: { [pic.id]: { text: 'Bar chart of students per region', source: 'Test model suggestion' } } });
  assert.ok(r.applied.some((a) => a.source === 'Test model suggestion'));
  assert.equal(r.result.summary.errors, 0);
});
