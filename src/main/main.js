'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, nativeImage } = require('electron');
const engine = require('../core/engine');
const { buildReport } = require('../core/report');
const settings = require('./settings');
const ai = require('./ai');

const docs = new Map();
let win = null;

function createWindow() {
  win = new BrowserWindow({
    width: 1280, height: 860, minWidth: 760, minHeight: 560,
    title: 'Accessible Document Checker',
    backgroundColor: '#F3F2F7',
    icon: path.join(__dirname, '..', 'renderer', 'icon.png'),
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true, spellcheck: true },
  });
  win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https:\/\//.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (!url.startsWith('file://')) { e.preventDefault(); if (/^https:\/\//.test(url)) shell.openExternal(url); } });
}

app.whenReady().then(() => {
  Menu.setApplicationMenu(buildMenu());
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

function buildMenu() {
  const isMac = process.platform === 'darwin';
  return Menu.buildFromTemplate([
    ...(isMac ? [{ role: 'appMenu' }] : []),
    { label: 'File', submenu: [
      { label: 'Add documents…', accelerator: 'CmdOrCtrl+O', click: () => win && win.webContents.send('menu', 'add') },
      { label: 'Settings…', accelerator: 'CmdOrCtrl+,', click: () => win && win.webContents.send('menu', 'settings') },
      { type: 'separator' }, isMac ? { role: 'close' } : { role: 'quit' },
    ] },
    { role: 'editMenu' },
    { label: 'View', submenu: [{ role: 'resetZoom' }, { role: 'zoomIn' }, { role: 'zoomOut' }, { type: 'separator' }, { role: 'togglefullscreen' }, ...(app.isPackaged ? [] : [{ role: 'toggleDevTools' }])] },
    { role: 'windowMenu' },
    { role: 'help', submenu: [{ label: 'About the checks and their limits', click: () => win && win.webContents.send('menu', 'help') }] },
  ]);
}

// ---------------------------------------------------------------------------
function opts() {
  const s = settings.get();
  return { defaultLanguage: s.defaultLanguage };
}
function view(doc) {
  return {
    id: doc.id, name: doc.name, path: doc.path, format: doc.format,
    result: doc.result, originalCount: doc.originalResult.issues.length,
    applied: doc.applied, canUndo: doc.history.length > 0, dirty: doc.dirty,
    savedPath: doc.savedPath, verification: doc.verification, humanChecks: doc.humanChecks,
  };
}
function getDoc(id) {
  const d = docs.get(id);
  if (!d) throw new Error('That document is no longer open.');
  return d;
}
function getIssue(doc, issueId) {
  const i = doc.result.issues.find((x) => x.id === issueId);
  if (!i) throw new Error('This issue has already been fixed or has changed. The list has been refreshed.');
  return i;
}
async function openPath(p) {
  const name = path.basename(p);
  const format = engine.detectFormat(name);
  const buf = fs.readFileSync(p);
  const result = await engine.audit(buf, name, opts());
  const doc = { id: crypto.randomUUID(), path: p, name, format, original: buf, working: buf, history: [], applied: [], originalResult: result, result, humanChecks: {}, savedPath: null, verification: null, dirty: false };
  docs.set(doc.id, doc);
  return view(doc);
}
function wrap(fn) {
  return async (_e, ...args) => {
    try { return { ok: true, value: await fn(...args) }; } catch (e) { return { ok: false, error: e.message || String(e) }; }
  };
}

ipcMain.handle('docs:pick', wrap(async () => {
  const r = await dialog.showOpenDialog(win, { title: 'Add documents', properties: ['openFile', 'multiSelections'], filters: [{ name: 'Documents', extensions: ['docx', 'pptx', 'pdf'] }] });
  return r.canceled ? [] : r.filePaths;
}));
ipcMain.handle('docs:open', wrap(async (paths) => {
  const out = [];
  for (const p of paths) {
    try { out.push({ ok: true, doc: await openPath(p) }); } catch (e) { out.push({ ok: false, name: path.basename(p), error: e.message }); }
  }
  return out;
}));
ipcMain.handle('doc:close', wrap(async (id) => { docs.delete(id); return true; }));
ipcMain.handle('doc:recheck', wrap(async (id) => {
  const doc = getDoc(id);
  doc.result = await engine.audit(doc.working, doc.name, opts());
  return view(doc);
}));

const BY = { auto: 'Automatic fix', input: 'Your text', ai: 'Text approved by you' };
ipcMain.handle('doc:fix', wrap(async (id, issueId, input = {}) => {
  const doc = getDoc(id);
  const issue = getIssue(doc, issueId);
  const next = await engine.applyFix(doc.working, doc.name, issue, input, opts());
  doc.history.push({ buffer: doc.working, appliedLen: doc.applied.length });
  doc.working = next;
  doc.dirty = true;
  const detail = input.decorative ? 'marked as decorative' : input.value ? `“${String(input.value).slice(0, 120)}”` : issue.fix.suggestion && issue.fix.kind === 'auto' ? `set to “${issue.fix.suggestion}”` : '';
  doc.applied.push({ title: issue.title, detail, location: issue.location.label, wcag: issue.wcag.map((c) => c.id), by: input.fromAI ? 'Claude suggestion, approved by you' : BY[issue.fix.kind] || 'Fix', at: new Date().toISOString() });
  doc.result = await engine.audit(doc.working, doc.name, opts());
  return view(doc);
}));
ipcMain.handle('doc:fixAll', wrap(async (id) => {
  const doc = getDoc(id);
  const r = await engine.fixAllAuto(doc.working, doc.name, opts(), (n, issue) => win && win.webContents.send('progress', { id, n, title: issue.title }));
  if (r.applied.length) {
    doc.history.push({ buffer: doc.working, appliedLen: doc.applied.length });
    doc.working = r.buffer;
    doc.dirty = true;
    for (const a of r.applied) doc.applied.push({ title: a.title, by: 'Automatic fix (Fix all)', at: new Date().toISOString() });
  }
  doc.result = r.result;
  return { doc: view(doc), applied: r.applied.length, errors: r.errors };
}));
ipcMain.handle('doc:undo', wrap(async (id) => {
  const doc = getDoc(id);
  const h = doc.history.pop();
  if (!h) return view(doc);
  doc.working = h.buffer;
  doc.applied = doc.applied.slice(0, h.appliedLen);
  doc.dirty = doc.history.length > 0;
  doc.result = await engine.audit(doc.working, doc.name, opts());
  return view(doc);
}));
ipcMain.handle('doc:humanCheck', wrap(async (id, checkId, status, note) => {
  const doc = getDoc(id);
  doc.humanChecks[checkId] = { status, note: note || '' };
  return view(doc);
}));
ipcMain.handle('doc:preview', wrap(async (id, issueId) => {
  const doc = getDoc(id);
  const issue = getIssue(doc, issueId);
  if (doc.format === 'pdf') return null;
  const img = await engine.getImage(doc.working, doc.name, issue.target);
  if (!img) return null;
  if (['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp', 'image/svg+xml'].includes(img.mime)) return `data:${img.mime};base64,${img.data.toString('base64')}`;
  const ni = nativeImage.createFromBuffer(img.data);
  return ni.isEmpty() ? null : ni.resize({ width: 480 }).toDataURL();
}));

ipcMain.handle('ai:suggest', wrap(async (id, issueId) => {
  const doc = getDoc(id);
  const issue = getIssue(doc, issueId);
  const s = settings.get();
  if (!s.aiEnabled) throw new Error('AI suggestions are turned off in Settings.');
  const apiKey = settings.getApiKey();
  const language = doc.result.meta.language || s.defaultLanguage;
  if (/^img-alt/.test(issue.rule)) {
    const image = await engine.getImage(doc.working, doc.name, issue.target);
    if (!image) throw new Error('The image could not be read for a suggestion. Please type the alt text yourself.');
    return ai.suggestAltText({ apiKey, model: s.model, image, context: issue.target.context, language, figureHint: doc.format === 'pdf' ? 'the figure (if there are several, the most prominent one without a caption)' : null });
  }
  if (issue.rule === 'slide-title') return ai.suggestText({ apiKey, model: s.model, kind: 'slide-title', context: issue.target.context, language });
  if (issue.rule === 'link-text') return ai.suggestText({ apiKey, model: s.model, kind: 'link', current: issue.target.current, url: issue.target.url, context: issue.target.context, language });
  throw new Error('No suggestion is available for this issue.');
}));

function uniquePath(p) {
  if (!fs.existsSync(p)) return p;
  const ext = path.extname(p), base = p.slice(0, -ext.length);
  for (let i = 2; i < 1000; i++) { const c = `${base} ${i}${ext}`; if (!fs.existsSync(c)) return c; }
  throw new Error('Could not find a free file name.');
}
ipcMain.handle('doc:save', wrap(async (id) => {
  const doc = getDoc(id);
  const s = settings.get();
  const dir = s.outputFolder && fs.existsSync(s.outputFolder) ? s.outputFolder : path.dirname(doc.path);
  let target = doc.savedPath && fs.existsSync(doc.savedPath) ? doc.savedPath : uniquePath(path.join(dir, engine.fixedName(doc.name)));
  if (path.resolve(target) === path.resolve(doc.path)) target = uniquePath(path.join(dir, engine.fixedName(doc.name)));
  fs.writeFileSync(target, doc.working);
  // Verify: read the saved copy back from disk and check it again.
  const saved = fs.readFileSync(target);
  doc.verification = await engine.audit(saved, path.basename(target), opts());
  doc.result = doc.verification;
  doc.savedPath = target;
  doc.dirty = false;
  return view(doc);
}));
ipcMain.handle('doc:saveAs', wrap(async (id) => {
  const doc = getDoc(id);
  const r = await dialog.showSaveDialog(win, { title: 'Save accessible copy', defaultPath: path.join(path.dirname(doc.path), engine.fixedName(doc.name)) });
  if (r.canceled || !r.filePath) return view(doc);
  if (path.resolve(r.filePath) === path.resolve(doc.path)) throw new Error('To protect your original, choose a different name or folder.');
  fs.writeFileSync(r.filePath, doc.working);
  doc.verification = await engine.audit(fs.readFileSync(r.filePath), path.basename(r.filePath), opts());
  doc.result = doc.verification;
  doc.savedPath = r.filePath;
  doc.dirty = false;
  return view(doc);
}));

ipcMain.handle('report:export', wrap(async (id, kind) => {
  const doc = getDoc(id);
  const html = buildReport({ fileName: doc.name, savedPath: doc.savedPath, format: doc.format, original: doc.originalResult, current: doc.result, applied: doc.applied, humanChecks: doc.humanChecks });
  const base = doc.name.replace(/\.[^.]+$/, '');
  const ext = kind === 'pdf' ? 'pdf' : kind === 'json' ? 'json' : 'html';
  const r = await dialog.showSaveDialog(win, { title: 'Export accessibility report', defaultPath: path.join(path.dirname(doc.savedPath || doc.path), `${base} - accessibility report.${ext}`), filters: [{ name: ext.toUpperCase(), extensions: [ext] }] });
  if (r.canceled || !r.filePath) return null;
  if (kind === 'json') {
    fs.writeFileSync(r.filePath, JSON.stringify({ file: doc.name, savedCopy: doc.savedPath, checkedAt: doc.result.checkedAt, summary: doc.result.summary, remainingIssues: doc.result.issues, fixesApplied: doc.applied, humanChecks: doc.result.humanChecks.map((h) => ({ ...h, status: (doc.humanChecks[h.id] || {}).status || 'not checked' })), originalIssues: doc.originalResult.issues }, null, 2));
  } else if (kind === 'pdf') {
    const tmp = path.join(os.tmpdir(), `adc-report-${Date.now()}.html`);
    fs.writeFileSync(tmp, html);
    const w = new BrowserWindow({ show: false, webPreferences: { sandbox: true, javascript: false } });
    try {
      await w.loadFile(tmp);
      const pdf = await w.webContents.printToPDF({ pageSize: 'A4', printBackground: true, generateTaggedPDF: true, generateDocumentOutline: true, margins: { marginType: 'default' } });
      fs.writeFileSync(r.filePath, pdf);
    } finally { w.destroy(); fs.rmSync(tmp, { force: true }); }
  } else fs.writeFileSync(r.filePath, html);
  return r.filePath;
}));

ipcMain.handle('settings:get', wrap(async () => settings.get()));
ipcMain.handle('settings:set', wrap(async (patch) => settings.set(patch)));
ipcMain.handle('settings:setKey', wrap(async (key) => settings.setApiKey((key || '').trim())));
ipcMain.handle('settings:testKey', wrap(async (key) => ai.testKey((key || '').trim() || settings.getApiKey(), settings.get().model)));
ipcMain.handle('settings:pickFolder', wrap(async () => {
  const r = await dialog.showOpenDialog(win, { title: 'Choose where to save accessible copies', properties: ['openDirectory', 'createDirectory'] });
  return r.canceled ? null : r.filePaths[0];
}));
ipcMain.handle('shell:show', wrap(async (p) => { if (p && fs.existsSync(p)) shell.showItemInFolder(p); return true; }));
ipcMain.handle('shell:open', wrap(async (p) => { if (p && fs.existsSync(p)) await shell.openPath(p); return true; }));
ipcMain.handle('shell:external', wrap(async (url) => { if (/^https:\/\//.test(url)) await shell.openExternal(url); return true; }));
