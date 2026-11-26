'use strict';
// Word (.docx) accessibility audit and remediation.
const JSZip = require('jszip');
const X = require('./xml');
const C = require('./contrast');
const { makeIssue, LINK_TEXT_BAD, isUrlText, suggestLinkText } = require('./issues');

const REL_NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const REL_CORE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CT_NUMBERING = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml';

const HIGHLIGHT = {
  yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
  darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
  darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF',
};
const LIST_MARKER = /^\s*([•·▪◦‣○■□➢►✓\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)])[\t ]+/;

// ---------------------------------------------------------------------------
// Loading
// ---------------------------------------------------------------------------
async function load(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const read = async (p) => (zip.file(p) ? X.parse(await zip.file(p).async('string')) : null);
  const ctx = {
    zip,
    doc: await read('word/document.xml'),
    styles: await read('word/styles.xml'),
    core: await read('docProps/core.xml'),
    rels: await read('word/_rels/document.xml.rels'),
    numbering: await read('word/numbering.xml'),
    dirty: new Set(),
  };
  if (!ctx.doc) throw new Error('This file does not look like a Word document (word/document.xml is missing).');
  ctx.styleMap = buildStyleMap(ctx.styles);
  return ctx;
}
async function save(ctx) {
  const parts = { doc: 'word/document.xml', styles: 'word/styles.xml', core: 'docProps/core.xml', rels: 'word/_rels/document.xml.rels', numbering: 'word/numbering.xml' };
  for (const key of ctx.dirty) {
    if (parts[key]) ctx.zip.file(parts[key], X.serialize(ctx[key]));
  }
  return ctx.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------
function buildStyleMap(styles) {
  const map = { byId: {}, defaults: { sz: 20, lang: null, color: null } };
  if (!styles) return map;
  const dd = X.first(styles, 'w', 'docDefaults');
  if (dd) {
    const rpr = X.first(dd, 'w', 'rPr');
    if (rpr) {
      const sz = X.kid(rpr, 'w', 'sz'); if (sz) map.defaults.sz = +X.attr(sz, 'w', 'val');
      const lang = X.kid(rpr, 'w', 'lang'); if (lang) map.defaults.lang = X.attr(lang, 'w', 'val');
      const color = X.kid(rpr, 'w', 'color'); if (color) map.defaults.color = X.attr(color, 'w', 'val');
    }
  }
  for (const s of X.kids(X.first(styles, 'w', 'styles') || styles.documentElement, 'w', 'style')) {
    const id = X.attr(s, 'w', 'styleId');
    const name = X.attr(X.kid(s, 'w', 'name'), 'w', 'val') || id;
    const basedOn = X.attr(X.kid(s, 'w', 'basedOn'), 'w', 'val');
    const rpr = X.kid(s, 'w', 'rPr');
    const ppr = X.kid(s, 'w', 'pPr');
    const outline = ppr ? X.kid(ppr, 'w', 'outlineLvl') : null;
    map.byId[id] = {
      id, name, basedOn, el: s, type: X.attr(s, null, 'w:type') || s.getAttributeNS(X.NS.w, 'type'),
      color: rpr && X.kid(rpr, 'w', 'color') ? X.attr(X.kid(rpr, 'w', 'color'), 'w', 'val') : null,
      sz: rpr && X.kid(rpr, 'w', 'sz') ? +X.attr(X.kid(rpr, 'w', 'sz'), 'w', 'val') : null,
      b: rpr && X.kid(rpr, 'w', 'b') ? X.attr(X.kid(rpr, 'w', 'b'), 'w', 'val') !== '0' && X.attr(X.kid(rpr, 'w', 'b'), 'w', 'val') !== 'false' : null,
      shd: ppr && X.kid(ppr, 'w', 'shd') ? X.attr(X.kid(ppr, 'w', 'shd'), 'w', 'fill') : null,
      outlineLvl: outline ? +X.attr(outline, 'w', 'val') : null,
    };
  }
  return map;
}
function styleChain(map, id) {
  const out = [];
  let cur = id, guard = 0;
  while (cur && map.byId[cur] && guard++ < 20) { out.push(map.byId[cur]); cur = map.byId[cur].basedOn; }
  return out;
}
function styleProp(map, id, key) {
  for (const s of styleChain(map, id)) if (s[key] !== null && s[key] !== undefined) return s[key];
  return null;
}
function headingLevelOfStyle(map, id) {
  if (!id) return 0;
  for (const s of styleChain(map, id)) {
    const n = (s.name || '').toLowerCase();
    const m = n.match(/^heading\s*([1-9])$/);
    if (m) return +m[1];
    if (n === 'title') return 1;
    if (s.outlineLvl !== null && s.outlineLvl < 9) return s.outlineLvl + 1;
  }
  return 0;
}
function pStyleOf(p) {
  const ppr = X.kid(p, 'w', 'pPr');
  return ppr ? X.attr(X.kid(ppr, 'w', 'pStyle'), 'w', 'val') : null;
}
function headingLevel(ctx, p) {
  const ppr = X.kid(p, 'w', 'pPr');
  const ol = ppr ? X.kid(ppr, 'w', 'outlineLvl') : null;
  if (ol && +X.attr(ol, 'w', 'val') < 9) return +X.attr(ol, 'w', 'val') + 1;
  return headingLevelOfStyle(ctx.styleMap, pStyleOf(p));
}
function isTitleStyle(ctx, p) {
  const id = pStyleOf(p);
  return id && ctx.styleMap.byId[id] && (ctx.styleMap.byId[id].name || '').toLowerCase() === 'title';
}
function findHeadingStyleId(ctx, level) {
  for (const s of Object.values(ctx.styleMap.byId)) {
    if ((s.name || '').toLowerCase().replace(/\s+/g, ' ') === `heading ${level}`) return s.id;
  }
  return null;
}
function ensureHeadingStyle(ctx, level) {
  const existing = findHeadingStyleId(ctx, level);
  if (existing) return existing;
  if (!ctx.styles) throw new Error('The document has no styles part, so a heading style cannot be added.');
  const d = ctx.styles;
  const sizes = { 1: 32, 2: 28, 3: 26, 4: 24, 5: 22, 6: 22 };
  const s = X.el(d, 'w', 'style', { 'w:type': 'paragraph', 'w:styleId': `Heading${level}` });
  const add = (parent, local, attrs) => { const e = X.el(d, 'w', local, attrs); parent.appendChild(e); return e; };
  add(s, 'name', { 'w:val': `heading ${level}` });
  add(s, 'basedOn', { 'w:val': 'Normal' });
  add(s, 'next', { 'w:val': 'Normal' });
  add(s, 'uiPriority', { 'w:val': '9' });
  add(s, 'qFormat');
  const ppr = add(s, 'pPr');
  add(ppr, 'keepNext'); add(ppr, 'spacing', { 'w:before': '240', 'w:after': '60' }); add(ppr, 'outlineLvl', { 'w:val': String(level - 1) });
  const rpr = add(s, 'rPr');
  add(rpr, 'b'); add(rpr, 'sz', { 'w:val': String(sizes[level] || 22) });
  (X.first(d, 'w', 'styles') || d.documentElement).appendChild(s);
  ctx.dirty.add('styles');
  ctx.styleMap = buildStyleMap(ctx.styles);
  return `Heading${level}`;
}
function setPStyle(p, styleId) {
  const ppr = X.ensureKid(p, 'w', 'pPr', true);
  let ps = X.kid(ppr, 'w', 'pStyle');
  if (!ps) { ps = X.el(p.ownerDocument, 'w', 'pStyle'); ppr.insertBefore(ps, ppr.firstChild); }
  ps.setAttributeNS(X.NS.w, 'w:val', styleId);
  const ol = X.kid(ppr, 'w', 'outlineLvl'); if (ol) X.remove(ol);
}

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------
function paraText(p) {
  let s = '';
  const walk = (n) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (c.namespaceURI === X.NS.w && c.localName === 't') s += c.textContent;
      else if (c.namespaceURI === X.NS.w && c.localName === 'tab') s += '\t';
      else if (c.namespaceURI === X.NS.w && (c.localName === 'br' || c.localName === 'cr')) s += ' ';
      else if (c.namespaceURI === X.NS.w && (c.localName === 'del' || c.localName === 'instrText')) continue;
      else if (c.namespaceURI === X.NS.w && c.localName === 'txbxContent') continue;
      else walk(c);
    }
  };
  walk(p);
  return s;
}
function short(s, n = 140) {
  s = (s || '').replace(/\s+/g, ' ').trim();
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}
function relTarget(ctx, rid) {
  if (!ctx.rels || !rid) return null;
  for (const r of X.all(ctx.rels, 'rel', 'Relationship')) {
    if (r.getAttribute('Id') === rid) return { target: r.getAttribute('Target'), external: r.getAttribute('TargetMode') === 'External', type: r.getAttribute('Type') };
  }
  return null;
}
function bodyParas(ctx) {
  return X.all(ctx.doc, 'w', 'p');
}
/** Map each paragraph to the text of the nearest preceding heading, for locations. */
function headingContext(ctx, paras) {
  const ctxMap = new Map();
  let current = null;
  for (const p of paras) {
    ctxMap.set(p, current); // a heading is located under the previous heading
    if (headingLevel(ctx, p) > 0 && paraText(p).trim()) current = short(paraText(p), 60);
  }
  return ctxMap;
}
function locate(ctx, paras, hctx, p, extra) {
  const i = paras.indexOf(p);
  const under = hctx.get(p);
  return {
    label: `Paragraph ${i + 1}${under ? ` (under “${under}”)` : ''}${extra ? ` · ${extra}` : ''}`,
    paragraph: i + 1,
    excerpt: short(paraText(p), 200),
  };
}

// ---------------------------------------------------------------------------
// Effective run formatting (for contrast)
// ---------------------------------------------------------------------------
function cellShading(p) {
  for (let n = p.parentNode; n; n = n.parentNode) {
    if (n.nodeType === 1 && n.namespaceURI === X.NS.w && n.localName === 'tc') {
      const tcPr = X.kid(n, 'w', 'tcPr');
      const shd = tcPr ? X.kid(tcPr, 'w', 'shd') : null;
      const fill = shd ? X.attr(shd, 'w', 'fill') : null;
      if (fill && /^[0-9A-Fa-f]{6}$/.test(fill)) return fill;
    }
  }
  return null;
}
function collectRuns(ctx) {
  const pageBgEl = X.first(ctx.doc, 'w', 'background');
  const pageBg = pageBgEl && /^[0-9A-Fa-f]{6}$/.test(X.attr(pageBgEl, 'w', 'color') || '') ? X.attr(pageBgEl, 'w', 'color') : 'FFFFFF';
  const out = [];
  for (const p of bodyParas(ctx)) {
    const ppr = X.kid(p, 'w', 'pPr');
    const pStyle = pStyleOf(p);
    const pShdEl = ppr ? X.kid(ppr, 'w', 'shd') : null;
    let bgPara = pShdEl ? X.attr(pShdEl, 'w', 'fill') : null;
    if (!bgPara || !/^[0-9A-Fa-f]{6}$/.test(bgPara)) bgPara = styleProp(ctx.styleMap, pStyle, 'shd');
    if (!bgPara || !/^[0-9A-Fa-f]{6}$/.test(bgPara)) bgPara = cellShading(p) || pageBg;
    for (const r of X.all(p, 'w', 'r')) {
      const t = X.all(r, 'w', 't').map((e) => e.textContent).join('');
      if (!t.trim()) continue;
      const rpr = X.kid(r, 'w', 'rPr');
      const rStyle = rpr ? X.attr(X.kid(rpr, 'w', 'rStyle'), 'w', 'val') : null;
      const colorEl = rpr ? X.kid(rpr, 'w', 'color') : null;
      let fg = colorEl ? X.attr(colorEl, 'w', 'val') : null;
      let source = colorEl ? 'run' : null, sourceStyle = null;
      if (!fg || fg === 'auto') {
        fg = styleProp(ctx.styleMap, rStyle, 'color'); if (fg) { source = 'style'; sourceStyle = styleChain(ctx.styleMap, rStyle).find((s) => s.color).id; }
      }
      if (!fg || fg === 'auto') {
        fg = styleProp(ctx.styleMap, pStyle, 'color'); if (fg) { source = 'style'; sourceStyle = styleChain(ctx.styleMap, pStyle).find((s) => s.color).id; }
      }
      if (!fg || fg === 'auto') { fg = ctx.styleMap.defaults.color; source = 'default'; }
      if (!fg || fg === 'auto' || !/^[0-9A-Fa-f]{6}$/.test(fg)) continue; // automatic colour adapts to the background
      let bg = bgPara;
      const rShd = rpr ? X.kid(rpr, 'w', 'shd') : null;
      if (rShd && /^[0-9A-Fa-f]{6}$/.test(X.attr(rShd, 'w', 'fill') || '')) bg = X.attr(rShd, 'w', 'fill');
      const hl = rpr ? X.kid(rpr, 'w', 'highlight') : null;
      if (hl && HIGHLIGHT[X.attr(hl, 'w', 'val')]) bg = HIGHLIGHT[X.attr(hl, 'w', 'val')];
      const szEl = rpr ? X.kid(rpr, 'w', 'sz') : null;
      const sz = (szEl ? +X.attr(szEl, 'w', 'val') : null) || styleProp(ctx.styleMap, rStyle, 'sz') || styleProp(ctx.styleMap, pStyle, 'sz') || ctx.styleMap.defaults.sz || 20;
      const bEl = rpr ? X.kid(rpr, 'w', 'b') : null;
      const bold = bEl ? !['0', 'false'].includes(X.attr(bEl, 'w', 'val') || '') : !!(styleProp(ctx.styleMap, rStyle, 'b') || styleProp(ctx.styleMap, pStyle, 'b'));
      out.push({ p, r, text: t, fg: fg.toUpperCase(), bg: bg.toUpperCase(), sizePt: sz / 2, bold, source, sourceStyle });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Audit
// ---------------------------------------------------------------------------
async function audit(buffer, opts = {}) {
  const ctx = await load(buffer);
  const issues = [];
  const paras = bodyParas(ctx);
  const hctx = headingContext(ctx, paras);
  const fileTitle = (opts.fileName || 'Document').replace(/\.[^.]+$/, '');

  // --- Title (2.4.2)
  const titleEl = ctx.core ? X.first(ctx.core, 'dc', 'title') : null;
  const title = X.text(titleEl).trim();
  let suggestedTitle = '';
  for (const p of paras) {
    if ((isTitleStyle(ctx, p) || headingLevel(ctx, p) === 1) && paraText(p).trim()) { suggestedTitle = short(paraText(p), 120); break; }
  }
  if (!suggestedTitle) suggestedTitle = fileTitle.replace(/[_-]+/g, ' ');
  if (!title) {
    issues.push(makeIssue({
      id: 'docx-title', rule: 'doc-title', wcag: ['2.4.2'],
      title: 'Document title is missing',
      description: 'Screen readers announce the document title (File > Info > Properties). Without it people hear the file name instead.',
      fix: { kind: 'auto', label: 'Set title', suggestion: suggestedTitle, editable: true, inputLabel: 'Document title' },
    }));
  }

  // --- Language (3.1.1)
  const lang = ctx.styleMap.defaults.lang;
  if (!lang) {
    issues.push(makeIssue({
      id: 'docx-lang', rule: 'doc-lang', wcag: ['3.1.1'],
      title: 'Document language is not set',
      description: 'Without a language, screen readers may read the text with the wrong pronunciation.',
      fix: { kind: 'auto', label: 'Set language', suggestion: opts.defaultLanguage || 'en-GB', editable: true, inputLabel: 'Language code (for example en-GB, fr-FR, ar-SA)' },
    }));
  }

  // --- Images and other graphics (1.1.1)
  X.all(ctx.doc, 'wp', 'docPr').forEach((dp, i) => {
    const container = dp.parentNode; // wp:inline or wp:anchor
    const descr = (dp.getAttribute('descr') || '').trim();
    const titleAttr = (dp.getAttribute('title') || '').trim();
    const decorative = X.isDecorative(dp);
    const blip = X.first(container, 'a', 'blip');
    const rid = blip ? X.attr(blip, 'r', 'embed') : null;
    const rel = relTarget(ctx, rid);
    const imagePart = rel && !rel.external ? 'word/' + rel.target.replace(/^\/?word\//, '').replace(/^\.\//, '') : null;
    const gd = X.first(container, 'a', 'graphicData');
    const uri = gd ? gd.getAttribute('uri') || '' : '';
    const kind = /chart/.test(uri) ? 'chart' : /diagram/.test(uri) ? 'SmartArt graphic' : /wordprocessingShape/.test(uri) ? 'shape' : 'image';
    if (kind === 'shape' && !blip) return; // plain shapes / text boxes are checked as floating content
    let p = container; while (p && !(p.namespaceURI === X.NS.w && p.localName === 'p')) p = p.parentNode;
    const loc = p ? locate(ctx, paras, hctx, p, `${kind} “${dp.getAttribute('name') || 'unnamed'}”`) : { label: kind };
    const ctxText = p ? contextAround(paras, p) : '';
    const looksLikeFileName = /\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(descr) || /^(picture|image|img|graphic|chart)[\s_-]*\d*$/i.test(descr) || /^IMG_\d+/i.test(descr);
    if (decorative) return;
    if (!descr && !titleAttr) {
      issues.push(makeIssue({
        id: `docx-img-${i}`, rule: 'img-alt', wcag: ['1.1.1'],
        title: `${cap(kind)} has no alt text`,
        description: `Add a short description of what this ${kind} shows or means, or mark it as decorative if it adds no information.`,
        location: loc,
        fix: { kind: imagePart ? 'ai' : 'input', label: 'Add alt text', inputLabel: 'Alt text', allowDecorative: true, multiline: true },
        target: { docPrIndex: i, imagePart, context: ctxText },
      }));
    } else if (looksLikeFileName) {
      issues.push(makeIssue({
        id: `docx-img-${i}`, rule: 'img-alt-quality', wcag: ['1.1.1'], severity: 'warning',
        title: `${cap(kind)} alt text looks like a file name`,
        description: `The alt text “${descr}” does not describe the ${kind}.`,
        location: loc,
        fix: { kind: imagePart ? 'ai' : 'input', label: 'Replace alt text', inputLabel: 'Alt text', allowDecorative: true, multiline: true, suggestion: '' },
        target: { docPrIndex: i, imagePart, context: ctxText },
      }));
    }
  });

  // --- Headings (1.3.1, 2.4.6)
  let prev = 0, headingCount = 0, textParas = 0;
  paras.forEach((p, i) => {
    const lvl = headingLevel(ctx, p);
    const t = paraText(p).trim();
    if (t) textParas++;
    if (!lvl || isTitleStyle(ctx, p)) { if (isTitleStyle(ctx, p) && t) prev = Math.max(prev, 0); return; }
    if (!t) {
      if (X.all(p, 'w', 'drawing').length === 0) {
        issues.push(makeIssue({
          id: `docx-emptyh-${i}`, rule: 'heading-empty', wcag: ['1.3.1', '2.4.6'], severity: 'warning',
          title: 'Empty heading', description: 'This paragraph uses a heading style but has no text, so screen reader users hear a blank heading.',
          location: locate(ctx, paras, hctx, p), fix: { kind: 'auto', label: 'Change to normal text' }, target: { pIndex: i },
        }));
      }
      return;
    }
    headingCount++;
    if (prev && lvl > prev + 1) {
      issues.push(makeIssue({
        id: `docx-hskip-${i}`, rule: 'heading-skip', wcag: ['1.3.1', '2.4.6'],
        title: `Heading level skipped (Heading ${prev} to Heading ${lvl})`,
        description: `Heading levels should go down one step at a time. This will be changed to Heading ${prev + 1}.`,
        location: locate(ctx, paras, hctx, p), fix: { kind: 'auto', label: `Make Heading ${prev + 1}` }, target: { pIndex: i, level: prev + 1 },
      }));
      prev = prev + 1;
    } else prev = lvl;
  });

  // Text that looks like a heading but isn't styled as one
  const fakeHeadings = [];
  paras.forEach((p, i) => {
    if (headingLevel(ctx, p) || isTitleStyle(ctx, p) || inTable(p)) return;
    const t = paraText(p).trim();
    if (!t || t.length > 90 || /[.,;:]$/.test(t) || LIST_MARKER.test(t)) return;
    const runs = X.all(p, 'w', 'r').filter((r) => X.all(r, 'w', 't').some((e) => e.textContent.trim()));
    if (!runs.length) return;
    const pStyle = pStyleOf(p);
    const allBold = runs.every((r) => { const rpr = X.kid(r, 'w', 'rPr'); const b = rpr && X.kid(rpr, 'w', 'b'); return (b && !['0', 'false'].includes(X.attr(b, 'w', 'val') || '')) || styleProp(ctx.styleMap, pStyle, 'b'); });
    const sizes = runs.map((r) => { const rpr = X.kid(r, 'w', 'rPr'); const s = rpr && X.kid(rpr, 'w', 'sz'); return s ? +X.attr(s, 'w', 'val') : (styleProp(ctx.styleMap, pStyle, 'sz') || ctx.styleMap.defaults.sz); });
    const big = Math.min(...sizes) >= Math.max(28, (ctx.styleMap.defaults.sz || 22) + 6);
    const next = paras[i + 1];
    const nextIsText = next && paraText(next).trim().length > t.length;
    if ((allBold || big) && nextIsText && t.split(/\s+/).length <= 12) fakeHeadings.push({ p, i, size: Math.min(...sizes) });
  });
  const sizesDesc = [...new Set(fakeHeadings.map((f) => f.size))].sort((a, b) => b - a);
  for (const f of fakeHeadings) {
    const hasRealH1 = paras.some((p) => headingLevel(ctx, p) === 1);
    let level = Math.min(3, sizesDesc.indexOf(f.size) + (hasRealH1 ? 2 : 1));
    issues.push(makeIssue({
      id: `docx-fakeh-${f.i}`, rule: 'heading-fake', wcag: ['1.3.1', '2.4.6'], severity: 'warning',
      title: 'Text looks like a heading but is not marked as one',
      description: 'Bold or large text is only a visual heading. Applying a real heading style lets screen reader users jump between sections.',
      location: locate(ctx, paras, hctx, f.p),
      fix: { kind: 'input', inputType: 'level', label: 'Make it a heading', inputLabel: 'Heading level', suggestion: String(level), inFixAll: false, dismissible: true },
      target: { pIndex: f.i },
    }));
  }
  if (!headingCount && textParas >= 12 && !fakeHeadings.length) {
    issues.push(makeIssue({
      id: 'docx-noheadings', rule: 'headings-none', wcag: ['1.3.1', '2.4.6'],
      title: 'Document has no headings',
      description: 'Long documents need headings so people can understand the structure and navigate. Apply Heading 1, Heading 2, etc. to section titles in Word.',
      fix: { kind: 'manual' },
    }));
  }

  // --- Tables (1.3.1)
  X.all(ctx.doc, 'w', 'tbl').forEach((tbl, i) => {
    const rows = X.kids(tbl, 'w', 'tr');
    if (!rows.length) return;
    const cols = Math.max(...rows.map((r) => X.kids(r, 'w', 'tc').length));
    const firstP = X.first(tbl, 'w', 'p');
    const loc = firstP ? { ...locate(ctx, paras, hctx, firstP) } : { label: `Table ${i + 1}` };
    loc.label = `Table ${i + 1} · ` + loc.label;
    if (rows.length < 2 || cols < 2) {
      issues.push(makeIssue({
        id: `docx-tbl-layout-${i}`, rule: 'table-layout', wcag: ['1.3.1', '1.3.2'], severity: 'warning',
        title: 'Possible layout table', description: 'This table has only one row or column. If it is used for layout, check the reading order makes sense; if it holds data, it needs a header row.',
        location: loc, fix: { kind: 'manual' },
      }));
      return;
    }
    const trPr = X.kid(rows[0], 'w', 'trPr');
    const hdr = trPr ? X.kid(trPr, 'w', 'tblHeader') : null;
    if (!hdr || ['0', 'false'].includes(X.attr(hdr, 'w', 'val') || '')) {
      issues.push(makeIssue({
        id: `docx-tbl-hdr-${i}`, rule: 'table-header', wcag: ['1.3.1'],
        title: 'Table has no header row',
        description: 'Marking the first row as a header row lets screen readers announce column headings for each cell (and repeats it across pages).',
        location: loc, fix: { kind: 'auto', label: 'Mark first row as header' }, target: { tblIndex: i },
      }));
    }
    const merged = X.all(tbl, 'w', 'gridSpan').some((g) => +X.attr(g, 'w', 'val') > 1) || X.all(tbl, 'w', 'vMerge').length > 0;
    const nested = X.all(tbl, 'w', 'tbl').length > 0;
    if (merged || nested) {
      issues.push(makeIssue({
        id: `docx-tbl-complex-${i}`, rule: 'table-complex', wcag: ['1.3.1'], severity: 'warning',
        title: nested ? 'Table contains a nested table' : 'Table has merged cells',
        description: 'Merged or nested cells are hard for screen readers to follow. Consider simplifying the table so each cell has one clear row and column heading.',
        location: loc, fix: { kind: 'manual' },
      }));
    }
  });

  // --- Contrast (1.4.3)
  const groups = new Map();
  for (const run of collectRuns(ctx)) {
    const need = C.requiredRatio(run.sizePt, run.bold);
    const r = C.ratio(run.fg, run.bg);
    if (r === null || r >= need) continue;
    const key = `${run.fg}|${run.bg}|${need}`;
    if (!groups.has(key)) groups.set(key, { ...run, need, ratio: r, count: 0 });
    groups.get(key).count++;
  }
  let gi = 0;
  for (const g of groups.values()) {
    const fixed = C.fixForeground(g.fg, g.bg, g.need);
    issues.push(makeIssue({
      id: `docx-contrast-${gi++}`, rule: 'contrast', wcag: ['1.4.3'],
      title: `Low colour contrast (${g.ratio.toFixed(2)}:1, needs ${g.need}:1)`,
      description: `Text colour #${g.fg} on #${g.bg} is hard to read${g.count > 1 ? ` (${g.count} places)` : ''}. Fixing changes the text colour to #${fixed}, keeping the same hue.`,
      location: { ...locate(ctx, paras, hctx, g.p), highlightText: g.text.trim().slice(0, 80) },
      fix: { kind: 'auto', label: `Change colour to #${fixed}` },
      target: { fg: g.fg, bg: g.bg, need: g.need, to: fixed, swatch: { fg: g.fg, bg: g.bg, to: fixed } },
    }));
  }

  // --- Manually typed lists (1.3.1)
  let run = [];
  const flush = () => {
    if (run.length >= 2) {
      const ordered = /^\s*\(?[0-9a-zA-Z]{1,3}[.)]/.test(paraText(run[0].p));
      issues.push(makeIssue({
        id: `docx-list-${run[0].i}`, rule: 'list-manual', wcag: ['1.3.1'],
        title: `List typed by hand (${run.length} items)`,
        description: 'These lines start with typed bullets or numbers, so screen readers do not announce them as a list. Fixing turns them into a real Word list.',
        location: locate(ctx, paras, hctx, run[0].p, `${run.length} items`),
        fix: { kind: 'auto', label: `Convert to ${ordered ? 'numbered' : 'bulleted'} list` },
        target: { pIndexes: run.map((r) => r.i), ordered },
      }));
    }
    run = [];
  };
  paras.forEach((p, i) => {
    const ppr = X.kid(p, 'w', 'pPr');
    const hasNum = ppr && X.kid(ppr, 'w', 'numPr');
    const t = paraText(p);
    if (!hasNum && LIST_MARKER.test(t) && t.replace(LIST_MARKER, '').trim()) {
      if (run.length && paras.indexOf(run[run.length - 1].p) !== i - 1) flush();
      run.push({ p, i });
    } else flush();
  });
  flush();

  // --- Links (2.4.4)
  X.all(ctx.doc, 'w', 'hyperlink').forEach((h, i) => {
    const t = paraText(h).trim();
    const rel = relTarget(ctx, X.attr(h, 'r', 'id'));
    const url = rel ? rel.target : (h.getAttributeNS(X.NS.w, 'anchor') ? '#' + h.getAttributeNS(X.NS.w, 'anchor') : '');
    if (!t) return;
    const bad = LINK_TEXT_BAD.test(t), raw = isUrlText(t);
    if (!bad && !raw) return;
    let p = h.parentNode; while (p && !(p.namespaceURI === X.NS.w && p.localName === 'p')) p = p.parentNode;
    issues.push(makeIssue({
      id: `docx-link-${i}`, rule: 'link-text', wcag: ['2.4.4'], severity: raw ? 'warning' : 'error',
      title: raw ? 'Link text is a raw web address' : `Link text “${t}” is not descriptive`,
      description: 'Link text should say where the link goes, so it makes sense when read on its own (screen reader users often list all links).',
      location: p ? { ...locate(ctx, paras, hctx, p), highlightText: t } : { label: 'Link' },
      fix: { kind: 'ai', label: 'Change link text', inputLabel: 'New link text', suggestion: suggestLinkText(url, t) },
      target: { linkIndex: i, url, current: t, context: p ? short(paraText(p), 400) : '' },
    }));
  });

  // --- Repeated blank paragraphs (1.3.1 best practice)
  const body = X.first(ctx.doc, 'w', 'body');
  if (body) {
    let blanks = [];
    const flushBlanks = () => {
      if (blanks.length >= 3) {
        issues.push(makeIssue({
          id: `docx-blank-${paras.indexOf(blanks[0])}`, rule: 'blank-paras', wcag: ['1.3.1'], severity: 'warning',
          title: `${blanks.length} empty paragraphs used for spacing`,
          description: 'Screen readers announce each empty line as “blank”. Use paragraph spacing or page breaks instead. Fixing keeps one empty line.',
          location: locate(ctx, paras, hctx, blanks[0]),
          fix: { kind: 'auto', label: 'Remove extra empty lines' }, target: { pIndexes: blanks.map((b) => paras.indexOf(b)) },
        }));
      }
      blanks = [];
    };
    for (const p of X.kids(body, 'w', 'p')) {
      const empty = !paraText(p).trim() && !X.all(p, 'w', 'drawing').length && !X.all(p, 'w', 'pict').length && !X.all(p, 'w', 'sectPr').length && !X.all(p, 'w', 'br').some((b) => X.attr(b, 'w', 'type') === 'page') && !X.all(p, 'w', 'bookmarkStart').length && !X.all(p, 'w', 'fldChar').length && !X.all(p, 'w', 'object').length;
      if (empty) blanks.push(p); else flushBlanks();
    }
    flushBlanks();
  }

  // --- Floating objects and text boxes (1.3.2)
  const anchors = X.all(ctx.doc, 'wp', 'anchor').length;
  const textBoxes = X.all(ctx.doc, 'w', 'txbxContent').length;
  if (anchors || textBoxes) {
    issues.push(makeIssue({
      id: 'docx-floating', rule: 'floating', wcag: ['1.3.2'], severity: 'warning',
      title: `${anchors + textBoxes} floating object${anchors + textBoxes > 1 ? 's' : ''} or text box${textBoxes > 1 ? 'es' : ''}`,
      description: 'Floating images and text boxes may be read out of order or skipped. Where possible set them to “In line with text” and move text out of text boxes.',
      fix: { kind: 'manual' },
    }));
  }

  return {
    format: 'docx',
    meta: { title: title || null, language: lang || null, paragraphs: paras.length, headings: headingCount, images: X.all(ctx.doc, 'wp', 'docPr').length, tables: X.all(ctx.doc, 'w', 'tbl').length },
    issues,
  };
}

function inTable(p) {
  for (let n = p.parentNode; n; n = n.parentNode) if (n.nodeType === 1 && n.namespaceURI === X.NS.w && n.localName === 'tc') return true;
  return false;
}
function contextAround(paras, p) {
  const i = paras.indexOf(p);
  const bits = [];
  for (let k = Math.max(0, i - 2); k <= Math.min(paras.length - 1, i + 2); k++) {
    const t = paraText(paras[k]).trim(); if (t) bits.push(t);
  }
  return short(bits.join('\n'), 800);
}
function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// ---------------------------------------------------------------------------
// Fixes
// ---------------------------------------------------------------------------
async function fix(buffer, issue, input = {}, opts = {}) {
  const ctx = await load(buffer);
  const t = issue.target || {};
  const paras = bodyParas(ctx);
  switch (issue.rule) {
    case 'doc-title': setTitle(ctx, (input.value || issue.fix.suggestion || '').trim()); break;
    case 'doc-lang': setLanguage(ctx, (input.value || issue.fix.suggestion || opts.defaultLanguage || 'en-GB').trim()); break;
    case 'img-alt':
    case 'img-alt-quality': {
      const dp = X.all(ctx.doc, 'wp', 'docPr')[t.docPrIndex];
      if (!dp) throw new Error('Image not found. Re-check the document and try again.');
      if (input.decorative) X.markDecorative(dp);
      else {
        const v = (input.value || '').trim();
        if (!v) throw new Error('Enter alt text, or mark the image as decorative.');
        dp.setAttribute('descr', v);
      }
      ctx.dirty.add('doc');
      break;
    }
    case 'heading-skip': {
      const p = paras[t.pIndex];
      setPStyle(p, ensureHeadingStyle(ctx, t.level));
      ctx.dirty.add('doc');
      break;
    }
    case 'heading-fake': {
      const level = Math.max(1, Math.min(6, parseInt(input.value || issue.fix.suggestion || '2', 10)));
      const p = paras[t.pIndex];
      setPStyle(p, ensureHeadingStyle(ctx, level));
      // Remove direct bold/size so the heading style controls the look.
      for (const r of X.all(p, 'w', 'r')) {
        const rpr = X.kid(r, 'w', 'rPr');
        if (rpr) for (const k of ['b', 'bCs', 'sz', 'szCs']) { const e = X.kid(rpr, 'w', k); if (e) X.remove(e); }
      }
      ctx.dirty.add('doc');
      break;
    }
    case 'heading-empty': {
      const p = paras[t.pIndex];
      const ppr = X.kid(p, 'w', 'pPr');
      const ps = ppr && X.kid(ppr, 'w', 'pStyle'); if (ps) X.remove(ps);
      const ol = ppr && X.kid(ppr, 'w', 'outlineLvl'); if (ol) X.remove(ol);
      ctx.dirty.add('doc');
      break;
    }
    case 'table-header': {
      const tbl = X.all(ctx.doc, 'w', 'tbl')[t.tblIndex];
      const row = X.kids(tbl, 'w', 'tr')[0];
      let trPr = X.kid(row, 'w', 'trPr');
      if (!trPr) {
        trPr = X.el(ctx.doc, 'w', 'trPr');
        const after = X.kid(row, 'w', 'tblPrEx');
        if (after) row.insertBefore(trPr, after.nextSibling); else row.insertBefore(trPr, row.firstChild);
      }
      let h = X.kid(trPr, 'w', 'tblHeader');
      if (!h) { h = X.el(ctx.doc, 'w', 'tblHeader'); trPr.appendChild(h); }
      h.removeAttributeNS(X.NS.w, 'val');
      ctx.dirty.add('doc');
      break;
    }
    case 'contrast': {
      for (const run of collectRuns(ctx)) {
        if (run.fg !== t.fg || run.bg !== t.bg || C.requiredRatio(run.sizePt, run.bold) !== t.need) continue;
        const rpr = X.ensureKid(run.r, 'w', 'rPr', true);
        let c = X.kid(rpr, 'w', 'color');
        if (!c) { c = X.el(ctx.doc, 'w', 'color'); rpr.appendChild(c); }
        c.setAttributeNS(X.NS.w, 'w:val', t.to);
        c.removeAttributeNS(X.NS.w, 'themeColor'); c.removeAttributeNS(X.NS.w, 'themeShade'); c.removeAttributeNS(X.NS.w, 'themeTint');
      }
      ctx.dirty.add('doc');
      break;
    }
    case 'link-text': {
      const h = X.all(ctx.doc, 'w', 'hyperlink')[t.linkIndex];
      const v = (input.value || '').trim();
      if (!v) throw new Error('Enter the new link text.');
      const ts = X.all(h, 'w', 't');
      ts.forEach((e, k) => { e.textContent = k === 0 ? v : ''; });
      if (ts[0]) ts[0].setAttribute('xml:space', 'preserve');
      ctx.dirty.add('doc');
      break;
    }
    case 'blank-paras': {
      t.pIndexes.slice(1).forEach((i) => X.remove(paras[i]));
      ctx.dirty.add('doc');
      break;
    }
    default:
      throw new Error('This issue cannot be fixed automatically.');
  }
  return save(ctx);
}

function setTitle(ctx, value) {
  if (!value) throw new Error('Enter a title.');
  if (!ctx.core) {
    ctx.core = X.parse(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${X.NS.cp}" xmlns:dc="${X.NS.dc}" xmlns:dcterms="${X.NS.dcterms}" xmlns:xsi="${X.NS.xsi}"></cp:coreProperties>`);
    ensurePackageCore(ctx);
  }
  let t = X.first(ctx.core, 'dc', 'title');
  if (!t) { t = ctx.core.createElementNS(X.NS.dc, 'dc:title'); ctx.core.documentElement.appendChild(t); }
  t.textContent = value;
  ctx.dirty.add('core');
}
function ensurePackageCore(ctx) {
  // Register docProps/core.xml in [Content_Types].xml and _rels/.rels.
  ctx.pendingCore = true;
  const origSave = ctx.zip.generateAsync.bind(ctx.zip);
  ctx.zip.generateAsync = async (o) => {
    const ct = X.parse(await ctx.zip.file('[Content_Types].xml').async('string'));
    if (!X.all(ct, 'ct', 'Override').some((e) => e.getAttribute('PartName') === '/docProps/core.xml')) {
      const ov = ct.createElementNS(X.NS.ct, 'Override'); ov.setAttribute('PartName', '/docProps/core.xml'); ov.setAttribute('ContentType', CT_CORE);
      ct.documentElement.appendChild(ov); ctx.zip.file('[Content_Types].xml', X.serialize(ct));
    }
    const rels = X.parse(await ctx.zip.file('_rels/.rels').async('string'));
    if (!X.all(rels, 'rel', 'Relationship').some((e) => e.getAttribute('Type') === REL_CORE)) {
      const r = rels.createElementNS(X.NS.rel, 'Relationship'); r.setAttribute('Id', 'rIdCore1'); r.setAttribute('Type', REL_CORE); r.setAttribute('Target', 'docProps/core.xml');
      rels.documentElement.appendChild(r); ctx.zip.file('_rels/.rels', X.serialize(rels));
    }
    return origSave(o);
  };
}
function setLanguage(ctx, lang) {
  if (!/^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$/.test(lang)) throw new Error('Use a language code such as en-GB.');
  if (!ctx.styles) throw new Error('The document has no styles part.');
  const d = ctx.styles;
  const root = X.first(d, 'w', 'styles') || d.documentElement;
  const dd = X.ensureKid(root, 'w', 'docDefaults', true);
  const rprd = X.ensureKid(dd, 'w', 'rPrDefault', true);
  const rpr = X.ensureKid(rprd, 'w', 'rPr');
  let l = X.kid(rpr, 'w', 'lang');
  if (!l) { l = X.el(d, 'w', 'lang'); rpr.appendChild(l); }
  l.setAttributeNS(X.NS.w, 'w:val', lang);
  ctx.dirty.add('styles');
}
async function ensureNumbering(ctx) {
  if (ctx.numbering) return;
  ctx.numbering = X.parse(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${X.NS.w}"></w:numbering>`);
  const ct = X.parse(await ctx.zip.file('[Content_Types].xml').async('string'));
  if (!X.all(ct, 'ct', 'Override').some((e) => e.getAttribute('PartName') === '/word/numbering.xml')) {
    const ov = ct.createElementNS(X.NS.ct, 'Override'); ov.setAttribute('PartName', '/word/numbering.xml'); ov.setAttribute('ContentType', CT_NUMBERING);
    ct.documentElement.appendChild(ov); ctx.zip.file('[Content_Types].xml', X.serialize(ct));
  }
  if (!X.all(ctx.rels, 'rel', 'Relationship').some((r) => r.getAttribute('Type') === REL_NUMBERING)) {
    const r = ctx.rels.createElementNS(X.NS.rel, 'Relationship');
    r.setAttribute('Id', 'rIdNumbering' + Date.now()); r.setAttribute('Type', REL_NUMBERING); r.setAttribute('Target', 'numbering.xml');
    ctx.rels.documentElement.appendChild(r); ctx.dirty.add('rels');
  }
}
// Converting typed lists needs async numbering-part setup, so it is handled here.
const _fix = fix;
async function fixWithLists(buffer, issue, input, opts) {
  if (issue.rule !== 'list-manual') return _fix(buffer, issue, input, opts);
  const ctx = await load(buffer);
  const paras = bodyParas(ctx);
  const ps = issue.target.pIndexes.map((i) => paras[i]);
  await ensureNumbering(ctx);
  const nd = ctx.numbering;
  const root = nd.documentElement;
  const absIds = X.kids(root, 'w', 'abstractNum').map((a) => +X.attr(a, 'w', 'abstractNumId'));
  const numIds = X.kids(root, 'w', 'num').map((n) => +X.attr(n, 'w', 'numId'));
  const absId = (absIds.length ? Math.max(...absIds) : 0) + 1;
  const numId = (numIds.length ? Math.max(...numIds) : 0) + 1;
  const ordered = issue.target.ordered;
  const abs = X.el(nd, 'w', 'abstractNum', { 'w:abstractNumId': String(absId) });
  abs.appendChild(X.el(nd, 'w', 'multiLevelType', { 'w:val': 'hybridMultilevel' }));
  for (let lvl = 0; lvl < 3; lvl++) {
    const l = X.el(nd, 'w', 'lvl', { 'w:ilvl': String(lvl) });
    l.appendChild(X.el(nd, 'w', 'start', { 'w:val': '1' }));
    l.appendChild(X.el(nd, 'w', 'numFmt', { 'w:val': ordered ? (lvl === 1 ? 'lowerLetter' : lvl === 2 ? 'lowerRoman' : 'decimal') : 'bullet' }));
    l.appendChild(X.el(nd, 'w', 'lvlText', { 'w:val': ordered ? `%${lvl + 1}.` : ['•', '◦', '▪'][lvl] }));
    l.appendChild(X.el(nd, 'w', 'lvlJc', { 'w:val': 'left' }));
    const ppr = X.el(nd, 'w', 'pPr');
    ppr.appendChild(X.el(nd, 'w', 'ind', { 'w:left': String(720 * (lvl + 1)), 'w:hanging': '360' }));
    l.appendChild(ppr);
    abs.appendChild(l);
  }
  const firstNum = X.kid(root, 'w', 'num');
  if (firstNum) root.insertBefore(abs, firstNum); else root.appendChild(abs);
  const num = X.el(nd, 'w', 'num', { 'w:numId': String(numId) });
  num.appendChild(X.el(nd, 'w', 'abstractNumId', { 'w:val': String(absId) }));
  root.appendChild(num);
  ctx.dirty.add('numbering');
  const listStyle = Object.values(ctx.styleMap.byId).find((s) => (s.name || '').toLowerCase() === 'list paragraph');
  for (const p of ps) {
    stripMarker(p);
    const ppr = X.ensureKid(p, 'w', 'pPr', true);
    if (listStyle && !X.kid(ppr, 'w', 'pStyle')) { const s = X.el(ctx.doc, 'w', 'pStyle', { 'w:val': listStyle.id }); ppr.insertBefore(s, ppr.firstChild); }
    const numPr = X.el(ctx.doc, 'w', 'numPr');
    numPr.appendChild(X.el(ctx.doc, 'w', 'ilvl', { 'w:val': '0' }));
    numPr.appendChild(X.el(ctx.doc, 'w', 'numId', { 'w:val': String(numId) }));
    const ps2 = X.kid(ppr, 'w', 'pStyle');
    const after = ps2 || null;
    // numPr follows pStyle, keepNext, keepLines, pageBreakBefore, framePr, widowControl
    let ref = after ? after.nextSibling : ppr.firstChild;
    while (ref && ref.nodeType === 1 && ['keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl'].includes(ref.localName)) ref = ref.nextSibling;
    ppr.insertBefore(numPr, ref);
    const ind = X.kid(ppr, 'w', 'ind'); if (ind) X.remove(ind);
  }
  ctx.dirty.add('doc');
  return save(ctx);
}
function stripMarker(p) {
  const m = paraText(p).match(LIST_MARKER);
  if (!m) return;
  let remaining = m[0].length;
  const nodes = [];
  const walk = (n) => { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; if (c.namespaceURI === X.NS.w && (c.localName === 't' || c.localName === 'tab')) nodes.push(c); else if (!(c.namespaceURI === X.NS.w && ['del', 'instrText', 'txbxContent'].includes(c.localName))) walk(c); } };
  walk(p);
  for (const n of nodes) {
    if (remaining <= 0) break;
    if (n.localName === 'tab') { X.remove(n); remaining -= 1; continue; }
    const s = n.textContent;
    if (s.length <= remaining) { remaining -= s.length; n.textContent = ''; } else { n.textContent = s.slice(remaining); n.setAttribute('xml:space', 'preserve'); remaining = 0; }
  }
}

/** Return the raw bytes of an image referenced by an issue. */
async function getImage(buffer, target) {
  if (!target || !target.imagePart) return null;
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(target.imagePart);
  if (!f) return null;
  return { data: await f.async('nodebuffer'), mime: mimeOf(target.imagePart) };
}
function mimeOf(name) {
  const ext = name.split('.').pop().toLowerCase();
  return { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf', svg: 'image/svg+xml' }[ext] || 'application/octet-stream';
}

module.exports = { audit, fix: fixWithLists, getImage, mimeOf };
