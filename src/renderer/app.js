'use strict';
/* global api */
// Renderer for Accessible Document Checker. Builds the UI with DOM APIs only
// (document content is never inserted as HTML).

const state = {
  docs: [],            // views from the main process
  currentId: null,
  view: 'dashboard',
  filter: 'all',
  drafts: {},          // docId -> issueId -> { value, decorative, fromAI }
  errors: {},          // docId -> issueId -> message
  busy: null,          // { label }
  settings: null,
  pendingFocus: null,
};

// ---------------------------------------------------------------------------
// Helpers
const PROVIDER_NAME = { anthropic: 'Claude', openai: 'OpenAI', ollama: 'the local model', off: '' };
function aiName() { const p = state.settings ? state.settings.provider : 'anthropic'; return PROVIDER_NAME[p] || 'AI'; }
function aiOn() { return !state.settings || state.settings.provider !== 'off'; }
/** Items "Fix all including suggested text" can fill in without asking an AI. */
function offlineSuggestible(i) { return ['ai', 'input'].includes(i.fix.kind) && i.fix.inFixAll !== false && i.fix.inputType !== 'level' && !!(i.fix.suggestion || '').trim(); }
function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  if (attrs) for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'text') el.textContent = v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (k === 'dataset') Object.assign(el.dataset, v);
    else if (v === true) el.setAttribute(k, '');
    else el.setAttribute(k, v);
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return el;
}
function svg(path, cls) {
  const ns = 'http://www.w3.org/2000/svg';
  const s = document.createElementNS(ns, 'svg');
  s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('aria-hidden', 'true'); s.setAttribute('focusable', 'false');
  if (cls) s.setAttribute('class', cls);
  const p = document.createElementNS(ns, 'path');
  p.setAttribute('d', path); p.setAttribute('fill', 'none'); p.setAttribute('stroke', 'currentColor'); p.setAttribute('stroke-width', '2.2'); p.setAttribute('stroke-linecap', 'round'); p.setAttribute('stroke-linejoin', 'round');
  s.append(p);
  return s;
}
const ICON = {
  check: 'M4 12.5l5 5L20 6.5',
  cross: 'M6 6l12 12M18 6L6 18',
  alert: 'M12 3l10 18H2L12 3zM12 10v5M12 18.5v.01',
  wand: 'M4 20L15 9M14 4v3M19 9h3M17.5 5.5l2-2M10 4.5l1 1M18.5 13.5l1 1',
  save: 'M5 3h11l3 3v15H5zM8 3v6h8V3M8 21v-7h8v7',
  refresh: 'M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7',
  undo: 'M9 14L4 9l5-5M4 9h10a6 6 0 0 1 0 12h-3',
  sparkle: 'M12 3l2 6 6 2-6 2-2 6-2-6-6-2 6-2z',
  doc: 'M6 2h9l5 5v15H6zM14 2v6h6',
  folder: 'M3 6h6l2 2h10v11H3z',
  report: 'M6 2h9l5 5v15H6zM9 13h6M9 17h6M9 9h2',
};
function announce(msg) {
  const live = document.getElementById('live');
  live.textContent = '';
  setTimeout(() => { live.textContent = msg; }, 50);
}
function toast(msg, isError) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast' + (isError ? ' error' : '');
  t.hidden = false;
  if (isError) { const a = document.getElementById('alert'); a.textContent = ''; setTimeout(() => { a.textContent = msg; }, 50); }
  else announce(msg);
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.hidden = true; }, isError ? 9000 : 5000);
}
const FORMAT_NAME = { docx: 'Word document', pptx: 'PowerPoint presentation', pdf: 'PDF' };
const FIX_KIND = { auto: 'Automatic fix', ai: 'Needs text: a suggestion is available', input: 'Needs your input', manual: 'Manual check or fix' };
function currentDoc() { return state.docs.find((d) => d.id === state.currentId) || null; }
function draft(docId, issueId) {
  state.drafts[docId] = state.drafts[docId] || {};
  return state.drafts[docId][issueId] || null;
}
function setDraft(docId, issueId, patch) {
  state.drafts[docId] = state.drafts[docId] || {};
  state.drafts[docId][issueId] = { ...(state.drafts[docId][issueId] || {}), ...patch };
}
function replaceDoc(view) {
  const i = state.docs.findIndex((d) => d.id === view.id);
  if (i >= 0) state.docs[i] = view; else state.docs.push(view);
}
async function busy(label, fn) {
  state.busy = { label };
  render();
  announce(label + '…');
  try { return await fn(); } finally { state.busy = null; render(); }
}
function outstandingHuman(doc) {
  return doc.result.humanChecks.filter((hc) => !['confirmed', 'not-applicable'].includes((doc.humanChecks[hc.id] || {}).status)).length;
}
function verdictOf(doc) {
  const s = doc.result.summary;
  const hc = outstandingHuman(doc);
  if (s.errors) return { cls: 'fail', icon: ICON.cross, title: `Not compliant yet: ${s.errors} error${s.errors > 1 ? 's' : ''} to fix`, text: `${s.warnings ? `${s.warnings} warning${s.warnings > 1 ? 's' : ''} to review. ` : ''}${s.auto ? `${s.auto} can be fixed automatically with “Fix all automatic”.` : ''}` };
  if (hc) return { cls: 'partial', icon: ICON.alert, title: 'Passes the automated checks', text: `${s.warnings ? `${s.warnings} warning${s.warnings > 1 ? 's' : ''} to review. ` : ''}${hc} human check${hc > 1 ? 's' : ''} still need a person to confirm before you can call it compliant.` };
  return { cls: 'pass', icon: ICON.check, title: 'Passes automated checks and all human checks are confirmed', text: s.warnings ? `${s.warnings} warning${s.warnings > 1 ? 's' : ''} remain as advice.` : 'No issues remain.' };
}

// ---------------------------------------------------------------------------
// Rendering
function render() {
  document.querySelectorAll('.nav-btn').forEach((b) => b.setAttribute('aria-current', b.dataset.view === state.view ? 'page' : 'false'));
  for (const v of ['dashboard', 'settings', 'help']) document.getElementById('view-' + v).hidden = state.view !== v;
  renderDocList();
  if (state.view === 'dashboard') renderDashboard();
  if (state.view === 'settings') renderSettings();
  if (state.view === 'help') renderHelp();
  if (state.pendingFocus) {
    const el = document.getElementById(state.pendingFocus) || document.querySelector(state.pendingFocus);
    state.pendingFocus = null;
    if (el) { el.focus(); el.scrollIntoView({ block: 'center' }); }
  }
}

function renderDocList() {
  const list = document.getElementById('doc-list');
  list.replaceChildren(...state.docs.map((d) => {
    const s = d.result.summary;
    const badge = s.errors ? h('span', { class: 'badge badge-error' }, `${s.errors} error${s.errors > 1 ? 's' : ''}`) : h('span', { class: 'badge badge-ok' }, 'Passes automated checks');
    return h('li', { class: 'doc-item' },
      h('button', { type: 'button', class: 'doc-btn', 'aria-current': d.id === state.currentId && state.view === 'dashboard' ? 'true' : 'false', onclick: () => { state.currentId = d.id; state.view = 'dashboard'; state.filter = 'all'; state.pendingFocus = 'doc-h1'; render(); } },
        h('span', { class: 'name' }, d.name),
        h('span', { class: 'small muted' }, FORMAT_NAME[d.format], d.dirty ? ' · unsaved fixes' : d.savedPath ? ' · saved' : ''),
        h('span', null, badge)),
      h('button', { type: 'button', class: 'doc-close', 'aria-label': `Close ${d.name}`, title: 'Close', onclick: () => closeDoc(d) }, '×'));
  }));
  document.getElementById('doc-count').textContent = state.docs.length ? `${state.docs.length} document${state.docs.length > 1 ? 's' : ''} open` : 'No documents yet.';
}

function renderDashboard() {
  const root = document.getElementById('view-dashboard');
  const doc = currentDoc();
  if (!doc) {
    root.setAttribute('aria-labelledby', 'welcome-heading');
    root.replaceChildren(h('div', { class: 'welcome' },
      h('h1', { id: 'welcome-heading' }, 'Check and fix document accessibility'),
      h('p', null, 'Add Word, PowerPoint or PDF files to check them against WCAG 2.1 level AA (and key PDF/UA rules for PDFs). Your files stay on this computer and fixes are always saved to a new copy, never over your original.'),
      h('ol', { class: 'steps' },
        h('li', null, h('strong', null, 'Add documents'), h('br'), 'Use “Add documents” or drag files onto the window.'),
        h('li', null, h('strong', null, 'Review the issues'), h('br'), 'Each issue shows where it is and which WCAG criterion it relates to.'),
        h('li', null, h('strong', null, 'Fix'), h('br'), 'Fix issues one by one, or use “Fix all”. Alt text, link text and slide titles can be suggested offline or by an AI model you choose, for you to review.'),
        h('li', null, h('strong', null, 'Save and verify'), h('br'), 'Saving writes an accessible copy and checks it again. Export a report for your records.')),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: addFiles }, svg('M12 5v14M5 12h14'), 'Add documents'),
        h('button', { type: 'button', class: 'btn', onclick: () => { state.view = 'help'; state.pendingFocus = 'help-heading'; render(); } }, 'What the checks cover'))));
    return;
  }
  root.setAttribute('aria-labelledby', 'doc-h1');
  const s = doc.result.summary;
  const v = verdictOf(doc);
  const children = [];
  children.push(h('div', { class: 'doc-header' },
    h('div', null,
      h('h1', { id: 'doc-h1', tabindex: '-1' }, doc.name),
      h('p', { class: 'muted' }, `${FORMAT_NAME[doc.format]} · last checked ${new Date(doc.result.checkedAt).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })} · ${doc.originalCount} issue${doc.originalCount === 1 ? '' : 's'} at first check, ${doc.applied.length} fix${doc.applied.length === 1 ? '' : 'es'} applied`))));
  children.push(h('div', { class: `verdict ${v.cls}`, role: 'status' }, svg(v.icon), h('div', null, h('h2', null, v.title), v.text ? h('p', null, v.text) : null)));
  if (doc.savedPath && doc.verification) {
    const ver = doc.verification.summary;
    children.push(h('section', { class: 'saved' + (ver.errors ? ' has-errors' : ''), 'aria-labelledby': 'saved-h' },
      h('h2', { id: 'saved-h' }, doc.dirty ? 'Accessible copy saved (newer fixes not saved yet)' : 'Accessible copy saved and re-checked'),
      h('p', null, 'Saved to ', h('strong', null, doc.savedPath), '. Your original file was not changed.'),
      h('p', null, ver.errors ? `Re-check of the saved file: ${ver.errors} error${ver.errors > 1 ? 's' : ''} remain.` : `Re-check of the saved file: passes all automated checks${ver.warnings ? ` (${ver.warnings} warning${ver.warnings > 1 ? 's' : ''} to review)` : ''}.`),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-small', onclick: () => api.openFile(doc.savedPath) }, svg(ICON.doc), 'Open the copy'),
        h('button', { type: 'button', class: 'btn btn-small', onclick: () => api.showInFolder(doc.savedPath) }, svg(ICON.folder), 'Show in folder'))));
  }
  children.push(h('ul', { class: 'tiles', 'aria-label': 'Issue counts' },
    h('li', { class: 'tile t-error' }, h('strong', null, s.errors), 'Errors'),
    h('li', { class: 'tile t-warn' }, h('strong', null, s.warnings), 'Warnings'),
    h('li', { class: 'tile t-auto' }, h('strong', null, s.auto), 'Can fix automatically'),
    h('li', { class: 'tile t-input' }, h('strong', null, s.ai + s.input), 'Need your text'),
    h('li', { class: 'tile' }, h('strong', null, s.manual), 'Manual'),
    h('li', { class: 'tile' }, h('strong', null, outstandingHuman(doc)), 'Human checks to do')));

  const autoCount = s.auto;
  const aiIssues = doc.result.issues.filter((i) => i.fix.kind === 'ai');
  const isBusy = !!state.busy;
  const offlineCount = doc.result.issues.filter(offlineSuggestible).length;
  const aiAltCount = aiOn() ? doc.result.issues.filter((i) => i.fix.kind === 'ai' && /^img-alt/.test(i.rule) && !(i.fix.suggestion || '').trim()).length : 0;
  const allCount = autoCount + offlineCount + aiAltCount;
  children.push(h('div', { class: 'toolbar', role: 'toolbar', 'aria-label': 'Document actions' },
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn btn-primary', disabled: isBusy || !autoCount, onclick: () => fixAll(doc) }, svg(ICON.wand), `Fix all automatic (${autoCount})`),
      h('button', { type: 'button', class: 'btn', disabled: isBusy || !allCount, 'aria-describedby': 'fixall-hint', onclick: () => fixAll(doc, true) }, svg(ICON.wand), `Fix all, including suggested text (${allCount})`),
      aiOn() ? h('button', { type: 'button', class: 'btn', disabled: isBusy || !aiIssues.length, onclick: () => suggestAll(doc) }, svg(ICON.sparkle), `Suggest text with ${aiName()} (${aiIssues.length})`) : null,
      h('button', { type: 'button', class: 'btn', disabled: isBusy, onclick: () => recheck(doc) }, svg(ICON.refresh), 'Re-check'),
      h('button', { type: 'button', class: 'btn', disabled: isBusy || !doc.canUndo, onclick: () => undo(doc) }, svg(ICON.undo), 'Undo last fix')),
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn btn-primary', disabled: isBusy, onclick: () => save(doc) }, svg(ICON.save), doc.savedPath && !doc.dirty ? 'Save again' : 'Save accessible copy'),
      h('label', { class: 'sr-only', for: 'export-kind' }, 'Report format'),
      h('select', { id: 'export-kind', class: 'btn-small', style: null }, h('option', { value: 'pdf' }, 'PDF report'), h('option', { value: 'html' }, 'HTML report'), h('option', { value: 'json' }, 'JSON data')),
      h('button', { type: 'button', class: 'btn', disabled: isBusy, onclick: () => exportReport(doc, document.getElementById('export-kind').value) }, svg(ICON.report), 'Export report'))));
  children.push(h('p', { class: 'muted small', id: 'fixall-hint' }, aiOn()
    ? `“Fix all, including suggested text” also fills in captions, chart data, link addresses and slide titles, and asks ${aiName()} for alt text where there is no caption. Every filled-in item is listed for you to review.`
    : '“Fix all, including suggested text” also fills in captions, chart data, link addresses and slide titles, without any AI. Every filled-in item is listed for you to review.'));
  const toReview = (doc.applied || []).filter((a) => a.suggested);
  if (toReview.length) {
    children.push(h('section', { class: 'panel review-panel', 'aria-labelledby': 'review-h' },
      h('h2', { id: 'review-h' }, `Suggested text to review (${toReview.length})`),
      h('p', null, 'These were filled in by “Fix all” from suggested text. Check each one is accurate. To change one, select Undo last fix, or edit it in the saved copy.'),
      h('ul', { class: 'review-list' }, toReview.map((a) => h('li', null, h('strong', null, a.title), a.location ? ` (${a.location})` : '', h('br'), a.detail || '', h('br'), h('span', { class: 'muted small' }, a.by))))));
  }
  if (state.busy) children.push(h('p', { class: 'muted', role: 'status' }, h('span', { class: 'spinner', 'aria-hidden': 'true' }), ' ', state.busy.label, '…'));

  // Filters
  const counts = {
    all: doc.result.issues.length,
    auto: doc.result.issues.filter((i) => i.fix.kind === 'auto').length,
    text: doc.result.issues.filter((i) => i.fix.kind === 'ai' || i.fix.kind === 'input').length,
    manual: doc.result.issues.filter((i) => i.fix.kind === 'manual').length,
    error: s.errors,
  };
  const filterBtn = (key, label) => h('button', { type: 'button', class: 'filter-btn', 'aria-pressed': state.filter === key ? 'true' : 'false', onclick: () => { state.filter = key; render(); } }, `${label} (${counts[key]})`);
  children.push(h('h2', { id: 'issues-h' }, 'Issues'));
  if (!doc.result.issues.length) {
    children.push(h('p', null, 'No automated issues found. Work through the human checks below to finish.'));
  } else {
    children.push(h('fieldset', { class: 'filters' }, h('legend', null, 'Show'),
      filterBtn('all', 'All'), filterBtn('error', 'Errors only'), filterBtn('auto', 'Automatic'), filterBtn('text', 'Needs your text'), filterBtn('manual', 'Manual')));
    const shown = doc.result.issues.filter((i) => state.filter === 'all' || (state.filter === 'error' && i.severity === 'error') || (state.filter === 'auto' && i.fix.kind === 'auto') || (state.filter === 'text' && (i.fix.kind === 'ai' || i.fix.kind === 'input')) || (state.filter === 'manual' && i.fix.kind === 'manual'));
    children.push(h('ol', { class: 'issues', 'aria-labelledby': 'issues-h' }, shown.map((i) => renderIssue(doc, i))));
  }
  children.push(renderHumanChecks(doc));
  root.replaceChildren(...children);
  loadPreviews(doc);
}

function highlightExcerpt(excerpt, needle) {
  if (!excerpt) return null;
  if (!needle) return h('p', { class: 'excerpt' }, excerpt);
  const i = excerpt.toLowerCase().indexOf(needle.toLowerCase());
  if (i < 0) return h('p', { class: 'excerpt' }, excerpt);
  return h('p', { class: 'excerpt' }, excerpt.slice(0, i), h('mark', null, excerpt.slice(i, i + needle.length)), excerpt.slice(i + needle.length));
}

function renderIssue(doc, issue) {
  const hid = `issue-${issue.id}`;
  const d = draft(doc.id, issue.id) || {};
  const err = (state.errors[doc.id] || {})[issue.id];
  const crit = issue.wcag.map((c) => `${c.id} ${c.name} (Level ${c.level})`).join('; ');
  const body = [];
  body.push(h('div', { class: 'issue-head' },
    h('h3', { id: hid, tabindex: '-1' }, issue.title),
    h('div', { class: 'badges' },
      h('span', { class: `badge ${issue.severity === 'error' ? 'badge-error' : 'badge-warn'}` }, issue.severity === 'error' ? 'Error' : 'Warning'),
      h('span', { class: 'badge badge-info' }, FIX_KIND[issue.fix.kind]))));
  body.push(h('p', { class: 'criteria' }, h('strong', null, 'WCAG 2.1: '), crit, issue.pdfua ? ` · PDF/UA ${issue.pdfua}` : ''));
  body.push(h('p', null, issue.description));
  body.push(h('p', { class: 'where' }, h('strong', null, 'Where: '), issue.location.label));
  const ex = highlightExcerpt(issue.location.excerpt, issue.location.highlightText);
  if (ex) body.push(ex);
  if (issue.target && issue.target.swatch) {
    const sw = issue.target.swatch;
    const chip = (fg, bg) => h('span', { class: 'chip', 'aria-hidden': 'true', dataset: { fg, bg } });
    body.push(h('div', { class: 'swatches' },
      h('span', { class: 'swatch' }, chip(sw.fg, sw.bg), `Now: text #${sw.fg} on #${sw.bg}`),
      h('span', { class: 'swatch' }, chip(sw.to, sw.bg), `After fix: text #${sw.to} on #${sw.bg}`)));
  }
  if (/^img-alt/.test(issue.rule) && doc.format !== 'pdf') body.push(h('img', { class: 'preview', alt: '', 'data-preview': issue.id, hidden: true }));
  body.push(renderFix(doc, issue, d, err));
  return h('li', { class: `issue sev-${issue.severity}` }, h('article', { 'aria-labelledby': hid }, body));
}

function renderFix(doc, issue, d, err) {
  const f = issue.fix;
  const box = h('div', { class: 'fix' });
  const errId = `err-${issue.id}`;
  if (err) box.append(h('p', { class: 'field-error', id: errId }, err));
  if (f.kind === 'manual') {
    box.append(h('p', { class: 'manual-note' }, 'This needs to be checked or fixed by a person, usually in the original program. It will stay listed so it is not forgotten.'));
    return box;
  }
  const inputId = `in-${issue.id}`;
  let control = null;
  if (f.inputType === 'level') {
    control = h('select', { id: inputId, onchange: (e) => setDraft(doc.id, issue.id, { value: e.target.value }) },
      [1, 2, 3, 4, 5, 6].map((n) => h('option', { value: String(n), selected: String(d.value || f.suggestion) === String(n) }, `Heading ${n}`)));
  } else if (f.kind !== 'auto' || f.editable) {
    const val = d.value !== undefined ? d.value : f.suggestion || '';
    const attrs = { id: inputId, 'aria-describedby': [err ? errId : null, `hint-${issue.id}`].filter(Boolean).join(' ') || null, oninput: (e) => setDraft(doc.id, issue.id, { value: e.target.value, fromAI: false }), disabled: d.decorative || null };
    control = f.multiline ? h('textarea', { ...attrs, rows: '3' }, val) : h('input', { ...attrs, type: 'text', value: val });
  }
  if (control) {
    const offline = f.suggestion && f.suggestionSource && d.value === undefined;
    const hint = d.fromAI ? `Suggested by ${typeof d.fromAI === 'string' ? d.fromAI : aiName()}. Check it is accurate and edit if needed before applying.`
      : offline ? `Suggested from the ${f.suggestionSource} (no AI used). Check it and edit if needed before applying.`
      : f.kind === 'ai' ? (aiOn() ? `Type the text, or ask ${aiName()} for a suggestion to review.` : 'Type the text.')
      : f.kind === 'auto' ? 'You can change this before fixing.' : f.suggestion ? 'A suggestion is filled in. Edit it if needed.' : '';
    box.append(h('div', { class: 'field' }, h('label', { for: inputId }, f.inputLabel || 'Value'), h('span', { class: 'hint', id: `hint-${issue.id}` }, hint), control));
  }
  if (f.allowDecorative) {
    const cid = `dec-${issue.id}`;
    box.append(h('div', { class: 'check' },
      h('input', { type: 'checkbox', id: cid, checked: !!d.decorative, onchange: (e) => { setDraft(doc.id, issue.id, { decorative: e.target.checked }); state.pendingFocus = cid; render(); } }),
      h('label', { for: cid }, 'This image is decorative (adds no information)')));
  }
  const row = h('div', { class: 'btn-row' });
  row.append(h('button', { type: 'button', class: 'btn btn-primary btn-small', disabled: !!state.busy, onclick: () => applyFix(doc, issue) }, svg(ICON.check), f.kind === 'auto' ? f.label || 'Fix' : `Apply${f.label ? ': ' + f.label.toLowerCase() : ''}`));
  if (f.kind === 'ai' && aiOn()) row.append(h('button', { type: 'button', class: 'btn btn-small', disabled: !!state.busy, onclick: () => suggestOne(doc, issue) }, svg(ICON.sparkle), d.fromAI ? 'Suggest again' : `Suggest with ${aiName()}`));
  box.append(row);
  return box;
}

async function loadPreviews(doc) {
  for (const img of document.querySelectorAll('img[data-preview]')) {
    const issueId = img.getAttribute('data-preview');
    try {
      const src = await api.preview(doc.id, issueId);
      if (src && img.isConnected) {
        img.src = src;
        img.hidden = false;
        img.alt = 'Preview of the image this issue is about';
      }
    } catch { /* preview is optional */ }
  }
  // Apply contrast swatch colours (set via JS so CSP can stay strict).
  for (const c of document.querySelectorAll('.chip[data-fg]')) { c.style.background = `linear-gradient(135deg, #${c.dataset.fg} 0 50%, #${c.dataset.bg} 50% 100%)`; }
}

function renderHumanChecks(doc) {
  const items = doc.result.humanChecks.map((hc) => {
    const st = (doc.humanChecks[hc.id] || {}).status || '';
    const name = `hc-${hc.id}`;
    const details = (doc.result.humanCheckDetails || {})[hc.id];
    const radio = (value, label) => h('label', { class: 'check' },
      h('input', { type: 'radio', name, value, checked: st === value, onchange: async () => { const v = await api.humanCheck(doc.id, hc.id, value); replaceDoc(v); state.pendingFocus = `#${CSS.escape(name)}-${value || 'none'}`; render(); announce(`${hc.title}: ${label}`); }, id: `${name}-${value || 'none'}` }), label);
    return h('li', { class: 'human-item' },
      h('h3', null, hc.title),
      h('p', { class: 'criteria' }, h('strong', null, 'WCAG 2.1: '), hc.wcag.map((c) => `${c.id} ${c.name} (Level ${c.level})`).join('; ')),
      h('p', null, hc.description),
      details ? h('details', null, h('summary', null, 'What to look at'), h('ul', null, details.map((x) => h('li', null, x)))) : null,
      h('fieldset', null, h('legend', null, `Status for “${hc.title}”`), radio('', 'Not checked yet'), radio('confirmed', 'I have checked this'), radio('not-applicable', 'Not applicable')));
  });
  return h('section', { class: 'human', 'aria-labelledby': 'human-h' },
    h('h2', { id: 'human-h' }, 'Human checks'),
    h('p', null, 'Software cannot judge these WCAG criteria. Confirm each one yourself; your answers are included in the report.'),
    h('ul', { class: 'human-list' }, items));
}

// ---------------------------------------------------------------------------
// Settings and help
async function renderSettings() {
  const root = document.getElementById('view-settings');
  if (!state.settings) state.settings = await api.getSettings();
  const s = state.settings;
  const prov = s.provider || 'anthropic';
  const save = async (patch, msg) => { state.settings = await api.setSettings(patch); if (msg) announce(msg); };
  const test = async () => {
    const el = document.getElementById('key-status');
    el.className = 'status-line'; el.textContent = 'Testing…';
    const keyEl = document.getElementById('api-key');
    const r = await api.testConnection({ apiKey: keyEl ? keyEl.value : '' }).catch((e) => ({ ok: false, message: e.message }));
    el.className = 'status-line ' + (r.ok ? 'ok' : 'bad'); el.textContent = r.ok ? `Connected. Using ${r.model}.` : r.message;
  };
  const radio = (value, label, hint) => h('div', { class: 'check radio' },
    h('input', { type: 'radio', name: 'provider', id: `prov-${value}`, value, checked: prov === value, 'aria-describedby': `prov-${value}-hint`, onchange: async () => { await save({ provider: value }, `${label} selected`); state.pendingFocus = `prov-${value}`; renderSettings(); } }),
    h('label', { for: `prov-${value}` }, label, h('span', { class: 'hint', id: `prov-${value}-hint` }, hint)));

  const keyBlock = (provider, label, hintText, placeholder, hasKey) => [
    h('div', { class: 'field' },
      h('label', { for: 'api-key' }, label),
      h('span', { class: 'hint', id: 'api-key-hint' }, `${hintText} It is stored ${s.keyStorage === 'encrypted' ? 'encrypted with your system keychain' : 'in your user settings folder (no system keychain was found)'} and only sent to that company.`),
      h('input', { type: 'password', id: 'api-key', autocomplete: 'off', spellcheck: 'false', placeholder: hasKey ? '•••••••• (saved, type to replace)' : placeholder, 'aria-describedby': 'api-key-hint key-status' })),
    h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn btn-primary', onclick: async () => { const v = document.getElementById('api-key').value; if (!v.trim()) { toast('Type a key first.', true); return; } state.settings = await api.setApiKey(v, provider); document.getElementById('api-key').value = ''; renderSettings(); toast('API key saved.'); } }, 'Save key'),
      h('button', { type: 'button', class: 'btn', onclick: test }, 'Test connection'),
      hasKey ? h('button', { type: 'button', class: 'btn btn-quiet', onclick: async () => { state.settings = await api.setApiKey('', provider); renderSettings(); toast('API key removed.'); } }, 'Remove key') : null),
    h('p', { id: 'key-status', class: 'status-line', role: 'status' }, hasKey ? 'An API key is saved on this computer.' : 'No API key saved yet.'),
  ];

  let details = [];
  if (prov === 'anthropic') {
    details = [...keyBlock('anthropic', 'Claude API key', 'Create a key at console.anthropic.com.', 'sk-ant-…', s.hasApiKey),
      h('div', { class: 'field' },
        h('label', { for: 'model' }, 'Claude model'),
        h('select', { id: 'model', onchange: (e) => save({ model: e.target.value }, 'Model saved') },
          [['claude-opus-5-5', 'Claude Opus 5.5 (best quality, default)'], ['claude-sonnet-5-5', 'Claude Sonnet 5.5 (faster, lower cost)'], ['claude-haiku-5-5', 'Claude Haiku 5.5 (fastest, lowest cost)']].map(([v, l]) => h('option', { value: v, selected: s.model === v }, l))))];
  } else if (prov === 'openai') {
    details = [...keyBlock('openai', 'OpenAI API key', 'Create a key at platform.openai.com.', 'sk-…', s.hasOpenaiKey),
      h('div', { class: 'field' },
        h('label', { for: 'openai-model' }, 'OpenAI model'),
        h('span', { class: 'hint', id: 'openai-model-hint' }, 'Any OpenAI model that accepts images, for example gpt-5 or gpt-5-mini.'),
        h('input', { type: 'text', id: 'openai-model', value: s.openaiModel, spellcheck: 'false', 'aria-describedby': 'openai-model-hint', onchange: (e) => save({ openaiModel: e.target.value.trim() || 'gpt-5' }, 'Model saved') }))];
  } else if (prov === 'ollama') {
    details = [
      h('div', { class: 'note' },
        h('p', null, h('strong', null, 'Set up once: '), 'install Ollama from ollama.com, open it, then download a vision model by running this in a terminal (about 3 GB):'),
        h('p', null, h('code', null, `ollama pull ${s.ollamaModel || 'gemma3'}`)),
        h('p', { class: 'muted small' }, 'Other vision models that work: qwen2.5vl, llama3.2-vision, granite3.2-vision. Bigger models describe images better but are slower. Suggestions take a few seconds to a minute depending on your computer.')),
      h('div', { class: 'field' },
        h('label', { for: 'ollama-model' }, 'Local model name'),
        h('input', { type: 'text', id: 'ollama-model', value: s.ollamaModel, spellcheck: 'false', onchange: async (e) => { await save({ ollamaModel: e.target.value.trim() || 'gemma3' }, 'Model saved'); } })),
      h('div', { class: 'field' },
        h('label', { for: 'ollama-url' }, 'Ollama address'),
        h('span', { class: 'hint', id: 'ollama-url-hint' }, 'Leave as http://localhost:11434 unless you run Ollama somewhere else.'),
        h('input', { type: 'url', id: 'ollama-url', value: s.ollamaUrl, spellcheck: 'false', 'aria-describedby': 'ollama-url-hint', onchange: (e) => save({ ollamaUrl: e.target.value.trim() || 'http://localhost:11434' }, 'Address saved') })),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: test }, 'Test connection')),
      h('p', { id: 'key-status', class: 'status-line', role: 'status' }, 'Not tested yet.')];
  }

  root.replaceChildren(
    h('h1', { id: 'settings-heading', tabindex: '-1' }, 'Settings'),
    h('section', { class: 'panel', 'aria-labelledby': 'ai-h' },
      h('h2', { id: 'ai-h' }, 'Suggestions for alt text, link text and slide titles'),
      h('p', null, 'Without any AI, the app already suggests text from captions, chart data, web addresses and slide text. To describe pictures that have no caption, choose a model below. You review every suggestion.'),
      h('fieldset', { class: 'form-grid' },
        h('legend', null, 'Who writes picture descriptions'),
        radio('ollama', 'A model on this computer (Ollama)', 'Free and fully offline. Nothing leaves your computer. Needs a one-off download.'),
        radio('anthropic', 'Claude, with an API key', 'Best quality. Uses the internet and is billed to your Anthropic account.'),
        radio('openai', 'OpenAI, with an API key', 'Uses the internet and is billed to your OpenAI account.'),
        radio('off', 'No AI', 'Only offline suggestions. Pictures without a caption need you to type the alt text.')),
      h('div', { class: 'form-grid' }, ...details),
      h('details', { class: 'note' },
        h('summary', null, 'Why can’t I sign in with my Claude or ChatGPT subscription?'),
        h('p', null, 'Anthropic’s terms do not allow other apps to use a Claude Pro or Max subscription; apps must use an API key. OpenAI’s “Sign in with ChatGPT” for using a Plus or Pro plan in other apps is currently open only to partners OpenAI approves. If that opens up, it can be added here.'))),
    h('section', { class: 'panel', 'aria-labelledby': 'doc-h' },
      h('h2', { id: 'doc-h' }, 'Documents'),
      h('div', { class: 'form-grid' },
        h('div', { class: 'field' },
          h('label', { for: 'lang' }, 'Default document language'),
          h('span', { class: 'hint', id: 'lang-hint' }, 'Used when a document has no language set. Use a code such as en-GB, en-US, fr-FR, es-ES, ar-SA or zh-CN.'),
          h('input', { type: 'text', id: 'lang', value: s.defaultLanguage, 'aria-describedby': 'lang-hint', onchange: async (e) => { if (!/^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$/.test(e.target.value.trim())) { toast('That does not look like a language code (for example en-GB).', true); return; } state.settings = await api.setSettings({ defaultLanguage: e.target.value.trim() }); announce('Default language saved'); } })),
        h('div', { class: 'field' },
          h('span', { class: 'label', id: 'out-label' }, 'Where to save accessible copies'),
          h('p', { id: 'out-value' }, s.outputFolder ? s.outputFolder : 'Next to the original file (named “… (accessible)”).'),
          h('div', { class: 'btn-row' },
            h('button', { type: 'button', class: 'btn', 'aria-describedby': 'out-label out-value', onclick: async () => { const p = await api.pickFolder(); if (p) { state.settings = await api.setSettings({ outputFolder: p }); renderSettings(); announce('Folder saved'); } } }, 'Choose folder…'),
            s.outputFolder ? h('button', { type: 'button', class: 'btn btn-quiet', onclick: async () => { state.settings = await api.setSettings({ outputFolder: '' }); renderSettings(); } }, 'Use the original’s folder') : null)))));
}

function renderHelp() {
  const root = document.getElementById('view-help');
  const row = (a, b, c) => h('tr', null, h('td', null, a), h('td', null, b), h('td', null, c));
  root.replaceChildren(
    h('h1', { id: 'help-heading', tabindex: '-1' }, 'About the checks'),
    h('section', { class: 'panel', 'aria-labelledby': 'h-tiers' },
      h('h2', { id: 'h-tiers' }, 'Three kinds of fix'),
      h('ul', { class: 'help-list' },
        h('li', null, h('strong', null, 'Automatic: '), 'the app can fix it safely without your input, for example document title and language, table header rows, heading levels, colour contrast, typed bullet lists, PDF tab order and bookmarks.'),
        h('li', null, h('strong', null, 'Needs your text: '), 'alt text, link text and slide titles need words that fit the meaning. The app suggests them offline where it can (from captions, chart data, web addresses and slide text), and an AI model you choose can suggest the rest. You review every suggestion.'),
        h('li', null, h('strong', null, 'Manual / human checks: '), 'some criteria can only be judged by a person, such as reading order, whether colour alone carries meaning, and whether alt text is accurate.'))),
    h('section', { class: 'panel', 'aria-labelledby': 'h-what' },
      h('h2', { id: 'h-what' }, 'What is checked'),
      h('table', { class: 'simple' },
        h('caption', { class: 'sr-only' }, 'Checks by document type'),
        h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Format'), h('th', { scope: 'col' }, 'Automated checks'), h('th', { scope: 'col' }, 'Not automated'))),
        h('tbody', null,
          row('Word', 'Title, language, image alt text, heading order and fake headings, table header rows, merged cells, colour contrast, typed lists, link text, repeated blank lines, floating objects', 'Reading order of text boxes, meaning conveyed by colour, quality of alt text'),
          row('PowerPoint', 'Title, language, slide titles (missing, duplicate, read first), picture alt text, table header rows, colour contrast of explicitly coloured text, typed lists, link text, media', 'Reading order of other objects, text that inherits template colours, captions for video'),
          row('PDF', 'Title and display title, language, tagged or not, figure alt text, heading order, table headers, link descriptions, tab order, bookmarks, font embedding, scanned pages', 'Creating tags for an untagged PDF, correctness of the tag tree and reading order, contrast')))),
    h('section', { class: 'panel', 'aria-labelledby': 'h-limits' },
      h('h2', { id: 'h-limits' }, 'Honest limits'),
      h('p', null, 'No tool can fully automate WCAG compliance. Passing every automated check is a strong start, but a document is only compliant once the human checks are also done. For PDFs, the best results come from fixing the original Word or PowerPoint file here and then exporting to PDF with accessibility tags turned on. For formal PDF/UA sign-off, also validate with a specialist tool such as PAC or veraPDF.'),
      h('p', null, 'This app does not add tags to untagged PDFs automatically, because automatic tagging is often wrong and can hide problems from checkers while still failing screen reader users.')),
    h('section', { class: 'panel', 'aria-labelledby': 'h-keys' },
      h('h2', { id: 'h-keys' }, 'Keyboard shortcuts'),
      h('ul', { class: 'help-list' },
        h('li', null, h('kbd', null, 'Ctrl/Cmd + O'), ' add documents'),
        h('li', null, h('kbd', null, 'Ctrl/Cmd + ,'), ' settings'),
        h('li', null, h('kbd', null, 'Ctrl/Cmd + plus / minus'), ' zoom in or out'))));
}

// ---------------------------------------------------------------------------
// Actions
async function addFiles() {
  try {
    const paths = await api.pickFiles();
    if (paths.length) await openPaths(paths);
  } catch (e) { toast(e.message, true); }
}
async function openPaths(paths) {
  await busy(`Checking ${paths.length} document${paths.length > 1 ? 's' : ''}`, async () => {
    const res = await api.openFiles(paths);
    const opened = res.filter((r) => r.ok).map((r) => r.doc);
    for (const d of opened) replaceDoc(d);
    const failed = res.filter((r) => !r.ok);
    if (opened.length) { state.currentId = opened[0].id; state.view = 'dashboard'; state.filter = 'all'; state.pendingFocus = 'doc-h1'; }
    if (failed.length) toast(failed.map((f) => `${f.name}: ${f.error}`).join(' '), true);
    else announce(`${opened.length} document${opened.length > 1 ? 's' : ''} checked. ${opened.map((d) => `${d.name}: ${d.result.summary.total} issues`).join('. ')}`);
  });
}
async function closeDoc(doc) {
  if (doc.dirty && !window.confirm(`${doc.name} has fixes that are not saved. Close it anyway?`)) return;
  await api.closeDoc(doc.id);
  state.docs = state.docs.filter((d) => d.id !== doc.id);
  if (state.currentId === doc.id) state.currentId = state.docs[0] ? state.docs[0].id : null;
  state.pendingFocus = 'add-btn';
  render();
  announce(`${doc.name} closed`);
}
function setError(doc, issue, msg) {
  state.errors[doc.id] = state.errors[doc.id] || {};
  if (msg) state.errors[doc.id][issue.id] = msg; else delete state.errors[doc.id][issue.id];
}
function nextFocusAfter(doc, issue) {
  const list = doc.result.issues;
  const idx = list.findIndex((i) => i.id === issue.id);
  return idx;
}
async function applyFix(doc, issue) {
  const d = draft(doc.id, issue.id) || {};
  const input = { decorative: !!d.decorative, fromAI: d.fromAI ? (typeof d.fromAI === 'string' ? d.fromAI : aiName()) : false, fromOffline: d.value === undefined && !!issue.fix.suggestionSource };
  if (issue.fix.inputType === 'level') input.value = d.value || issue.fix.suggestion;
  else if (issue.fix.kind !== 'auto' || issue.fix.editable) input.value = d.value !== undefined ? d.value : issue.fix.suggestion || '';
  if ((issue.fix.kind === 'ai' || issue.fix.kind === 'input') && !input.decorative && !String(input.value || '').trim()) {
    setError(doc, issue, 'Enter the text first' + (issue.fix.kind === 'ai' && aiOn() ? `, or ask ${aiName()} for a suggestion.` : '.'));
    state.pendingFocus = `in-${issue.id}`;
    render();
    return;
  }
  const idx = nextFocusAfter(doc, issue);
  try {
    const v = await busy(`Fixing: ${issue.title}`, () => api.fix(doc.id, issue.id, input));
    setError(doc, issue, null);
    if (state.drafts[doc.id]) delete state.drafts[doc.id][issue.id];
    replaceDoc(v);
    const remaining = v.result.issues;
    const still = remaining.find((i) => i.id === issue.id && i.title === issue.title);
    const next = still || remaining[Math.min(idx, remaining.length - 1)];
    state.pendingFocus = next ? `issue-${next.id}` : 'doc-h1';
    render();
    announce(`Fixed: ${issue.title}. ${remaining.length} issue${remaining.length === 1 ? '' : 's'} left.`);
  } catch (e) {
    setError(doc, issue, e.message);
    state.pendingFocus = `issue-${issue.id}`;
    render();
    toast(e.message, true);
  }
}
async function fixAll(doc, includeSuggestions = false) {
  try {
    const label = includeSuggestions ? 'Applying automatic fixes and suggested text' : 'Applying automatic fixes';
    const r = await busy(label, () => api.fixAll(doc.id, { includeSuggestions, useAI: includeSuggestions && aiOn() }));
    replaceDoc(r.doc);
    state.pendingFocus = r.suggested ? 'review-h' : 'doc-h1';
    render();
    const left = r.doc.result.summary;
    const auto = r.applied - (r.suggested || 0);
    toast(`${auto} automatic fix${auto === 1 ? '' : 'es'} applied.${r.suggested ? ` ${r.suggested} item${r.suggested === 1 ? '' : 's'} filled in from suggested text: please review them.` : ''}${r.errors.length ? ` ${r.errors.length} could not be done (${r.errors[0].message})` : ''} ${left.total} issue${left.total === 1 ? '' : 's'} left. Remember to save the accessible copy.`, r.errors.length > 0);
  } catch (e) { toast(e.message, true); }
}
async function suggestOne(doc, issue, quiet) {
  try {
    const r = await (quiet ? api.suggest(doc.id, issue.id) : busy(`Asking ${aiName()} about: ${issue.title}`, () => api.suggest(doc.id, issue.id)));
    setDraft(doc.id, issue.id, { value: r.text, decorative: !!r.decorative, fromAI: r.provider || aiName() });
    setError(doc, issue, null);
    if (!quiet) { state.pendingFocus = r.decorative ? `dec-${issue.id}` : `in-${issue.id}`; render(); announce(r.decorative ? `${r.provider || aiName()} thinks this image is decorative. Review and apply.` : `Suggestion ready: ${r.text}. Review it, then apply.`); }
    return true;
  } catch (e) {
    setError(doc, issue, e.message);
    if (!quiet) { render(); toast(e.message, true); }
    return false;
  }
}
async function suggestAll(doc) {
  const list = doc.result.issues.filter((i) => i.fix.kind === 'ai' && !(draft(doc.id, i.id) || {}).fromAI);
  if (!list.length) { toast('Every item already has a suggestion to review.'); return; }
  let ok = 0;
  await busy(`Asking ${aiName()} for ${list.length} suggestion${list.length > 1 ? 's' : ''}`, async () => {
    for (const i of list) { if (await suggestOne(doc, i, true)) ok++; }
  });
  state.filter = 'text';
  const first = doc.result.issues.find((i) => i.fix.kind === 'ai');
  state.pendingFocus = first ? `issue-${first.id}` : null;
  render();
  toast(`${ok} of ${list.length} suggestions ready. Review each one and select Apply. Nothing has been changed yet.`, ok < list.length);
}
async function recheck(doc) {
  try { const v = await busy('Re-checking', () => api.recheck(doc.id)); replaceDoc(v); render(); announce(`Re-checked. ${v.result.summary.total} issues.`); } catch (e) { toast(e.message, true); }
}
async function undo(doc) {
  try { const v = await busy('Undoing', () => api.undo(doc.id)); replaceDoc(v); render(); announce('Last fix undone.'); } catch (e) { toast(e.message, true); }
}
async function save(doc) {
  try {
    const v = await busy('Saving the accessible copy and re-checking it', () => api.save(doc.id));
    replaceDoc(v);
    state.pendingFocus = 'saved-h';
    render();
    const s = v.verification.summary;
    toast(`Saved. ${s.errors ? `${s.errors} error${s.errors > 1 ? 's' : ''} remain.` : 'The saved copy passes the automated checks.'}`, false);
  } catch (e) { toast(e.message, true); }
}
async function exportReport(doc, kind) {
  try {
    const p = await busy('Exporting the report', () => api.exportReport(doc.id, kind));
    if (p) toast(`Report saved to ${p}`);
  } catch (e) { toast(e.message, true); }
}

// ---------------------------------------------------------------------------
// Wiring
document.getElementById('add-btn').addEventListener('click', addFiles);
document.querySelectorAll('.nav-btn').forEach((b) => b.addEventListener('click', () => {
  state.view = b.dataset.view;
  state.pendingFocus = { dashboard: currentDoc() ? 'doc-h1' : 'welcome-heading', settings: 'settings-heading', help: 'help-heading' }[state.view];
  render();
}));
const dz = document.getElementById('dropzone');
for (const t of [document.body]) {
  t.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('over'); });
  t.addEventListener('dragleave', (e) => { if (!e.relatedTarget) dz.classList.remove('over'); });
  t.addEventListener('drop', (e) => {
    e.preventDefault();
    dz.classList.remove('over');
    const paths = [...e.dataTransfer.files].map((f) => api.pathForFile(f)).filter(Boolean);
    if (paths.length) openPaths(paths);
  });
}
api.onMenu((what) => {
  if (what === 'add') addFiles();
  else { state.view = what === 'help' ? 'help' : 'settings'; state.pendingFocus = state.view + '-heading'; render(); }
});
api.onProgress((p) => { if (state.busy) { state.busy.label = /^Getting a suggestion/.test(p.title) ? p.title : `Applying fixes (${p.n} done: ${p.title})`; const el = document.querySelector('main [role="status"].muted'); if (el) el.lastChild.textContent = state.busy.label + '…'; } });

api.getSettings().then((s) => { state.settings = s; render(); }).catch(() => render());
