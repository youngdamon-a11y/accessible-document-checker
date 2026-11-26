// End-to-end smoke test: launches the real Electron app, opens the sample files,
// fixes them, saves accessible copies and runs axe-core against the app's own UI.
// Run: xvfb-run -a node test/e2e.js [screenshotDir]
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const { _electron: electron } = require('playwright');

// A stand-in for Ollama, so the local-model route can be tested without a 3 GB download.
function fakeOllama() {
  const seen = { chats: 0, images: 0 };
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', (c) => { body += c; });
    req.on('end', () => {
      res.setHeader('content-type', 'application/json');
      if (req.url === '/api/tags') return res.end(JSON.stringify({ models: [{ name: 'gemma3:latest' }] }));
      if (req.url === '/api/chat') {
        const j = JSON.parse(body);
        seen.chats++;
        const imgs = (j.messages.find((m) => m.images) || {}).images || [];
        seen.images += imgs.length;
        return res.end(JSON.stringify({ message: { role: 'assistant', content: imgs.length ? 'Bar chart of students per region in 2026; the fourth region has the most students.' : 'Our results this year' } }));
      }
      res.statusCode = 404; res.end('{}');
    });
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ server, seen, url: `http://127.0.0.1:${server.address().port}` })));
}

(async () => {
  const shots = process.argv[2] || fs.mkdtempSync(path.join(os.tmpdir(), 'adc-shots-'));
  fs.mkdirSync(shots, { recursive: true });
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'adc-e2e-'));
  for (const f of fs.readdirSync(path.join(__dirname, 'fixtures'))) fs.copyFileSync(path.join(__dirname, 'fixtures', f), path.join(work, f));
  const ollama = await fakeOllama();
  const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'adc-userdata-'));
  const app = await electron.launch({ executablePath: require('electron'), args: [path.join(__dirname, '..'), '--no-sandbox'], env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: '1', ADC_USER_DATA: userData } });
  const page = await app.firstWindow();
  await page.setViewportSize({ width: 1280, height: 900 });
  page.on('console', (m) => { if (m.type() === 'error') console.log('console error:', m.text()); });
  page.on('pageerror', (e) => console.log('page error:', e.message));
  const axe = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
  const runAxe = async (label) => {
    await page.evaluate(axe);
    const r = await page.evaluate(() => window.axe.run(document, { runOnly: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'best-practice'] }));
    console.log(`axe [${label}]: ${r.violations.length} violations`);
    for (const v of r.violations) console.log(`  - ${v.id} (${v.impact}): ${v.help} — ${v.nodes.length} node(s): ${v.nodes.slice(0, 2).map((n) => n.target.join(' ')).join(' | ')}`);
    return r.violations.length;
  };
  let failures = 0;
  await page.waitForSelector('#welcome-heading');
  await page.screenshot({ path: path.join(shots, '1-welcome.png') });
  failures += await runAxe('welcome');

  const files = ['sample-inaccessible.docx', 'sample-inaccessible.pptx', 'sample-untagged.pdf'].map((f) => path.join(work, f));
  await page.evaluate((p) => openPaths(p), files);
  await page.waitForSelector('#doc-h1');
  await page.screenshot({ path: path.join(shots, '2-docx-issues.png'), fullPage: false });
  failures += await runAxe('docx issues');
  await page.locator('#issues-h').scrollIntoViewIfNeeded();
  await page.evaluate(() => document.getElementById('issues-h').scrollIntoView());
  await page.screenshot({ path: path.join(shots, '2b-docx-issue-cards.png') });

  // Fix all automatic, then provide alt text by hand, then save.
  await page.getByRole('button', { name: /Fix all automatic/ }).click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  const afterFixAll = await page.textContent('.verdict');
  console.log('after fix all:', afterFixAll.replace(/\s+/g, ' '));
  const alt = page.getByLabel('Alt text', { exact: true }).first();
  await alt.fill('Bar chart of students per region in 2026: four regions, the fourth has the most students.');
  await page.getByRole('button', { name: /Apply: add alt text/ }).first().click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  const link = page.getByLabel('New link text', { exact: true }).first();
  await link.fill('English programme data 2026');
  await page.getByRole('button', { name: /Apply: change link text/ }).first().click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  await page.getByRole('button', { name: /Save accessible copy/ }).click();
  await page.waitForSelector('#saved-h');
  await page.screenshot({ path: path.join(shots, '3-docx-saved.png') });
  const verdict = (await page.textContent('.verdict')).replace(/\s+/g, ' ');
  console.log('docx verdict after save:', verdict);
  if (!/Passes the automated checks/.test(verdict)) { failures++; console.log('FAIL: docx should pass automated checks'); }
  failures += await runAxe('saved');

  // Human checks + report export (dialog stubbed)
  await page.getByLabel('I have checked this').first().check();
  const reportPath = path.join(work, 'report.pdf');
  await app.evaluate(({ dialog }, p) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: p }); }, reportPath);
  await page.selectOption('#export-kind', 'pdf');
  await page.getByRole('button', { name: 'Export report' }).click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  for (let i = 0; i < 20 && !fs.existsSync(reportPath); i++) await new Promise((r) => setTimeout(r, 250));
  console.log('report pdf exists:', fs.existsSync(reportPath), fs.existsSync(reportPath) ? fs.statSync(reportPath).size : 0);
  if (!fs.existsSync(reportPath)) failures++;

  // Settings: choose the local model (fake Ollama) and test the connection
  await page.getByRole('button', { name: 'Settings' }).click();
  await page.waitForSelector('#settings-heading');
  await page.getByLabel(/A model on this computer/).check();
  await page.waitForSelector('#ollama-url');
  await page.fill('#ollama-url', ollama.url);
  await page.press('#ollama-url', 'Tab');
  await page.waitForTimeout(200);
  await page.getByRole('button', { name: 'Test connection' }).click();
  await page.waitForFunction(() => /Connected|not|Could/.test(document.getElementById('key-status').textContent));
  const conn = await page.textContent('#key-status');
  console.log('ollama test:', conn);
  if (!/Connected/.test(conn)) { failures++; console.log('FAIL: local model connection'); }
  await page.screenshot({ path: path.join(shots, '6-settings.png') });
  failures += await runAxe('settings (local model)');

  // PowerPoint: Fix all, including suggested text (chart data, slide text, link address, local model)
  await page.getByRole('button', { name: 'Dashboard' }).click();
  await page.getByRole('button', { name: /sample-inaccessible\.pptx/ }).first().click();
  await page.waitForSelector('#doc-h1');
  await page.getByRole('button', { name: /Fix all, including suggested text/ }).click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  await page.waitForSelector('#review-h');
  const reviewed = await page.$$eval('.review-list li', (els) => els.map((e) => e.textContent.replace(/\s+/g, ' ')));
  console.log('pptx suggested items:', reviewed.length); reviewed.forEach((r) => console.log('  -', r.slice(0, 160)));
  console.log('fake ollama calls:', ollama.seen);
  if (!ollama.seen.images) { failures++; console.log('FAIL: local model was not sent an image'); }
  const pv = (await page.textContent('.verdict')).replace(/\s+/g, ' ');
  console.log('pptx verdict:', pv);
  if (!/Passes the automated checks/.test(pv)) { failures++; console.log('FAIL: pptx should pass automated checks after Fix all with suggestions'); }
  await page.screenshot({ path: path.join(shots, '4-pptx.png') });
  failures += await runAxe('pptx');

  // PDF page rendering for local models
  const png = await app.evaluate(async (_e, b64) => (await global.__adcRenderPdf(Buffer.from(b64, 'base64'))).subarray(0, 8).toString('hex'), fs.readFileSync(path.join(work, 'sample-untagged.pdf')).toString('base64'));
  console.log('pdf page render header:', png);
  if (png !== '89504e470d0a1a0a') { failures++; console.log('FAIL: PDF page did not render to PNG'); }

  // PDF
  await page.getByRole('button', { name: /sample-untagged\.pdf/ }).first().click();
  await page.waitForSelector('#doc-h1');
  await page.getByRole('button', { name: /Fix all automatic/ }).click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  await page.screenshot({ path: path.join(shots, '5-pdf.png') });
  console.log('pdf verdict:', (await page.textContent('.verdict')).replace(/\s+/g, ' '));

  // Settings and help views
  await page.getByRole('button', { name: 'Settings' }).click();
  await page.waitForSelector('#settings-heading');
  await page.getByLabel(/Claude, with an API key/).check();
  await page.waitForSelector('#api-key');
  await page.screenshot({ path: path.join(shots, '6b-settings-claude.png') });
  failures += await runAxe('settings (Claude)');
  await page.getByRole('button', { name: 'About the checks' }).click();
  await page.waitForSelector('#help-heading');
  failures += await runAxe('help');

  // Original files must be untouched
  for (const f of ['sample-inaccessible.docx', 'sample-inaccessible.pptx', 'sample-untagged.pdf']) {
    const same = Buffer.compare(fs.readFileSync(path.join(work, f)), fs.readFileSync(path.join(__dirname, 'fixtures', f))) === 0;
    if (!same) { failures++; console.log('FAIL: original modified', f); }
  }
  console.log('files in work dir:', fs.readdirSync(work).join(', '));
  await app.close();
  ollama.server.close();
  console.log(failures ? `E2E finished with ${failures} problem(s)` : 'E2E passed');
  console.log('screenshots:', shots);
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
