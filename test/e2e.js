// End-to-end smoke test: launches the real Electron app, opens the sample files,
// fixes them, saves accessible copies and runs axe-core against the app's own UI.
// Run: xvfb-run -a node test/e2e.js [screenshotDir]
const fs = require('fs');
const os = require('os');
const path = require('path');
const { _electron: electron } = require('playwright');

(async () => {
  const shots = process.argv[2] || fs.mkdtempSync(path.join(os.tmpdir(), 'adc-shots-'));
  fs.mkdirSync(shots, { recursive: true });
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'adc-e2e-'));
  for (const f of fs.readdirSync(path.join(__dirname, 'fixtures'))) fs.copyFileSync(path.join(__dirname, 'fixtures', f), path.join(work, f));
  const app = await electron.launch({ executablePath: require('electron'), args: [path.join(__dirname, '..'), '--no-sandbox'], env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: '1' } });
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

  // PowerPoint
  await page.getByRole('button', { name: /sample-inaccessible\.pptx/ }).first().click();
  await page.waitForSelector('#doc-h1');
  await page.getByRole('button', { name: /Fix all automatic/ }).click();
  await page.waitForFunction(() => !document.querySelector('.spinner'));
  await page.screenshot({ path: path.join(shots, '4-pptx.png') });
  failures += await runAxe('pptx');

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
  await page.screenshot({ path: path.join(shots, '6-settings.png') });
  failures += await runAxe('settings');
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
  console.log(failures ? `E2E finished with ${failures} problem(s)` : 'E2E passed');
  console.log('screenshots:', shots);
  process.exit(failures ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
