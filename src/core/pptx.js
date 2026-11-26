'use strict';
// PowerPoint (.pptx) accessibility audit and remediation.
const path = require('path').posix;
const JSZip = require('jszip');
const X = require('./xml');
const C = require('./contrast');
const { makeIssue, LINK_TEXT_BAD, isUrlText, suggestLinkText } = require('./issues');
const D = require('./describe');
const { mimeOf } = require('./docx');

const REL_CORE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml';
const LIST_MARKER = /^\s*([•·▪◦‣○■□➢►✓\-–—*]|\(?\d{1,3}[.)]|\(?[a-zA-Z][.)])[\t ]+/;
const SCHEME_ALIAS = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };

// ---------------------------------------------------------------------------
async function load(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const read = async (p) => (zip.file(p) ? X.parse(await zip.file(p).async('string')) : null);
  const pres = await read('ppt/presentation.xml');
  if (!pres) throw new Error('This file does not look like a PowerPoint presentation (ppt/presentation.xml is missing).');
  const presRels = await read('ppt/_rels/presentation.xml.rels');
  const relMap = (rels) => {
    const m = {};
    if (rels) for (const r of X.all(rels, 'rel', 'Relationship')) m[r.getAttribute('Id')] = { target: r.getAttribute('Target'), type: r.getAttribute('Type'), external: r.getAttribute('TargetMode') === 'External', el: r };
    return m;
  };
  const pr = relMap(presRels);
  const slides = [];
  for (const sid of X.all(pres, 'p', 'sldId')) {
    const rid = X.attr(sid, 'r', 'id');
    if (!pr[rid]) continue;
    const p = path.normalize(path.join('ppt', pr[rid].target.replace(/^\//, '')));
    const relsPath = path.join(path.dirname(p), '_rels', path.basename(p) + '.rels');
    const rels = await read(relsPath);
    slides.push({ path: p, doc: await read(p), relsPath, rels, relMap: relMap(rels), number: slides.length + 1, dirty: false });
  }
  // Theme colours
  let theme = null;
  const themeRel = Object.values(pr).find((r) => /\/theme$/.test(r.type));
  if (themeRel) theme = await read(path.normalize(path.join('ppt', themeRel.target)));
  if (!theme) { const tf = Object.keys(zip.files).find((f) => /^ppt\/theme\/theme\d+\.xml$/.test(f)); if (tf) theme = await read(tf); }
  const scheme = {};
  if (theme) {
    const cs = X.first(theme, 'a', 'clrScheme');
    if (cs) for (let c = cs.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const srgb = X.first(c, 'a', 'srgbClr'); const sys = X.first(c, 'a', 'sysClr');
      scheme[c.localName] = srgb ? srgb.getAttribute('val') : sys ? sys.getAttribute('lastClr') : null;
    }
  }
  // Layout/master backgrounds for each slide
  for (const s of slides) {
    s.bgChain = [];
    const layoutRel = Object.values(s.relMap).find((r) => /\/slideLayout$/.test(r.type));
    if (layoutRel) {
      const lp = path.normalize(path.join(path.dirname(s.path), layoutRel.target));
      const layout = await read(lp);
      if (layout) s.bgChain.push(layout);
      const lrels = await read(path.join(path.dirname(lp), '_rels', path.basename(lp) + '.rels'));
      const mrel = lrels && X.all(lrels, 'rel', 'Relationship').find((r) => /\/slideMaster$/.test(r.getAttribute('Type')));
      if (mrel) { const master = await read(path.normalize(path.join(path.dirname(lp), mrel.getAttribute('Target')))); if (master) s.bgChain.push(master); }
    }
  }
  return { zip, pres, presRels, core: await read('docProps/core.xml'), slides, scheme, dirty: new Set() };
}
async function save(ctx) {
  if (ctx.dirty.has('pres')) ctx.zip.file('ppt/presentation.xml', X.serialize(ctx.pres));
  if (ctx.dirty.has('core')) {
    ctx.zip.file('docProps/core.xml', X.serialize(ctx.core));
    if (ctx.newCore) await registerCore(ctx);
  }
  for (const s of ctx.slides) if (s.dirty) ctx.zip.file(s.path, X.serialize(s.doc));
  return ctx.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
async function registerCore(ctx) {
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
}

// ---------------------------------------------------------------------------
function shapes(slide) {
  const tree = X.first(slide.doc, 'p', 'spTree');
  if (!tree) return [];
  const out = [];
  for (let c = tree.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1 || c.namespaceURI !== X.NS.p) continue;
    if (['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(c.localName)) out.push(c);
  }
  return out;
}
function cNvPrOf(shape) {
  return X.first(shape, 'p', 'cNvPr');
}
function phType(shape) {
  const ph = X.first(shape, 'p', 'ph');
  if (!ph) return null;
  return ph.getAttribute('type') || 'body';
}
function isTitleShape(shape) {
  const t = phType(shape);
  return shape.localName === 'sp' && (t === 'title' || t === 'ctrTitle');
}
function shapeText(shape) {
  const body = X.first(shape, 'p', 'txBody');
  if (!body) return '';
  return X.all(body, 'a', 'p').map((p) => X.all(p, 'a', 't').map((t) => t.textContent).join('')).join('\n');
}
function slideTitle(slide) {
  const t = shapes(slide).find(isTitleShape);
  return t ? shapeText(t).replace(/\s+/g, ' ').trim() : '';
}
function slideText(slide) {
  return shapes(slide).map(shapeText).filter(Boolean).join('\n').slice(0, 800);
}
function short(s, n = 140) {
  s = (s || '').replace(/\s+/g, ' ').trim();
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}
function describeShape(shape) {
  const c = cNvPrOf(shape);
  const name = c ? c.getAttribute('name') : 'shape';
  if (isTitleShape(shape)) return `Title “${short(shapeText(shape), 40) || name}”`;
  if (shape.localName === 'pic') return `Picture “${name}”`;
  if (shape.localName === 'graphicFrame') return `${X.first(shape, 'a', 'tbl') ? 'Table' : 'Chart/graphic'} “${name}”`;
  const t = short(shapeText(shape), 40);
  return t ? `Text “${t}”` : `Shape “${name}”`;
}
function loc(slide, extra) {
  const title = slideTitle(slide);
  return { label: `Slide ${slide.number}${title ? ` (“${short(title, 50)}”)` : ''}${extra ? ` · ${extra}` : ''}`, slide: slide.number, excerpt: short(title || slideText(slide), 200) };
}

// Colours ------------------------------------------------------------------
function colourOf(fillParent, scheme) {
  if (!fillParent) return null;
  const solid = X.kid(fillParent, 'a', 'solidFill');
  return solid ? colourFromClr(solid, scheme) : null;
}
/** Resolve the colour element (srgbClr / schemeClr) directly inside `container`. */
function colourFromClr(container, scheme) {
  const srgb = X.kid(container, 'a', 'srgbClr');
  if (srgb) return hasMods(srgb) ? null : srgb.getAttribute('val').toUpperCase();
  const sc = X.kid(container, 'a', 'schemeClr');
  if (sc && !hasMods(sc)) { const v = sc.getAttribute('val'); const hex = scheme[SCHEME_ALIAS[v] || v]; return hex ? hex.toUpperCase() : null; }
  return null;
}
function hasMods(clr) {
  for (let c = clr.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.localName !== 'alpha') return true;
  return false;
}
function slideBg(slide, scheme) {
  for (const d of [slide.doc, ...slide.bgChain]) {
    const bgPr = X.first(d, 'p', 'bgPr');
    if (bgPr) { const c = colourOf(bgPr, scheme); if (c) return c; if (X.kid(bgPr, 'a', 'blipFill') || X.kid(bgPr, 'a', 'gradFill')) return null; }
    const bgRef = X.first(d, 'p', 'bgRef');
    if (bgRef) { const c = colourFromClr(bgRef, scheme); if (c) return c; }
  }
  return scheme.lt1 ? scheme.lt1.toUpperCase() : 'FFFFFF';
}
function textRuns(ctx, slide) {
  const out = [];
  const bg = slideBg(slide, ctx.scheme);
  for (const shape of shapes(slide)) {
    const spPr = X.first(shape, 'p', 'spPr');
    let shapeBg = spPr ? colourOf(spPr, ctx.scheme) : null;
    if (spPr && !shapeBg && (X.kid(spPr, 'a', 'blipFill') || X.kid(spPr, 'a', 'gradFill'))) continue; // picture/gradient behind text: cannot judge
    const body = X.first(shape, 'p', 'txBody');
    if (!body) continue;
    // Tables: cells have their own fill
    for (const r of X.all(shape, 'a', 'r')) {
      const t = X.text(X.kid(r, 'a', 't'));
      if (!t.trim()) continue;
      const rPr = X.kid(r, 'a', 'rPr');
      const fg = rPr ? colourOf(rPr, ctx.scheme) : null;
      if (!fg) continue; // inherited colours come from the template; checked by the human contrast review
      let back = shapeBg || bg;
      const tc = ancestor(r, 'a', 'tc');
      if (tc) { const tcPr = X.kid(tc, 'a', 'tcPr'); const cb = tcPr ? colourOf(tcPr, ctx.scheme) : null; if (cb) back = cb; }
      if (!back) continue;
      const sz = rPr && rPr.getAttribute('sz') ? +rPr.getAttribute('sz') / 100 : null;
      const bold = rPr && rPr.getAttribute('b') === '1';
      out.push({ r, shape, text: t, fg, bg: back, sizePt: sz, bold });
    }
  }
  return out;
}
function ancestor(n, ns, local) {
  for (let p = n.parentNode; p; p = p.parentNode) if (p.nodeType === 1 && p.namespaceURI === X.NS[ns] && p.localName === local) return p;
  return null;
}

// ---------------------------------------------------------------------------
async function audit(buffer, opts = {}) {
  const ctx = await load(buffer);
  const issues = [];
  const fileTitle = (opts.fileName || 'Presentation').replace(/\.[^.]+$/, '');

  // Title (2.4.2)
  const title = X.text(ctx.core ? X.first(ctx.core, 'dc', 'title') : null).trim();
  if (!title) {
    const suggestion = (ctx.slides[0] && slideTitle(ctx.slides[0])) || fileTitle.replace(/[_-]+/g, ' ');
    issues.push(makeIssue({
      id: 'pptx-title', rule: 'doc-title', wcag: ['2.4.2'], title: 'Presentation title is missing',
      description: 'The title in File > Info > Properties is announced by screen readers and shown when exported to PDF.',
      fix: { kind: 'auto', label: 'Set title', suggestion, editable: true, inputLabel: 'Presentation title' },
    }));
  }

  // Language (3.1.1)
  let runsTotal = 0, runsNoLang = 0;
  for (const s of ctx.slides) for (const r of X.all(s.doc, 'a', 'r')) { runsTotal++; const rp = X.kid(r, 'a', 'rPr'); if (!rp || !rp.getAttribute('lang')) runsNoLang++; }
  const defR = X.first(ctx.pres, 'a', 'defRPr');
  const presLang = defR ? defR.getAttribute('lang') : '';
  if (!presLang && (runsTotal === 0 || runsNoLang / runsTotal > 0.5)) {
    issues.push(makeIssue({
      id: 'pptx-lang', rule: 'doc-lang', wcag: ['3.1.1'], title: 'Presentation language is not set',
      description: 'Without a language, screen readers may use the wrong voice and pronunciation.',
      fix: { kind: 'auto', label: 'Set language', suggestion: opts.defaultLanguage || 'en-GB', editable: true, inputLabel: 'Language code (for example en-GB)' },
    }));
  }

  const seenTitles = new Map();
  const readingOrder = [];
  for (const slide of ctx.slides) {
    const shs = shapes(slide);
    // Slide titles (2.4.2 / 2.4.6)
    const titleShape = shs.find(isTitleShape);
    const st = titleShape ? shapeText(titleShape).replace(/\s+/g, ' ').trim() : '';
    if (!st) {
      const firstText = shs.map(shapeText).find((t) => t.trim()) || '';
      issues.push(makeIssue({
        id: `pptx-notitle-${slide.number}`, rule: 'slide-title', wcag: ['2.4.2', '2.4.6'],
        title: `Slide ${slide.number} has no title`,
        description: titleShape
          ? 'The title placeholder is empty. Screen reader users navigate by slide titles.'
          : 'Screen reader users navigate by slide titles. A title will be added just above the visible slide area, so your design does not change but the title is still announced.',
        location: loc(slide),
        fix: { kind: 'ai', label: 'Add slide title', inputLabel: 'Slide title', suggestion: short(firstText.split('\n')[0], 80) },
        target: { slide: slide.number - 1, hasPlaceholder: !!titleShape, context: slideText(slide) },
      }));
    } else {
      const key = st.toLowerCase();
      if (seenTitles.has(key)) {
        issues.push(makeIssue({
          id: `pptx-duptitle-${slide.number}`, rule: 'slide-title-dup', wcag: ['2.4.6'], severity: 'warning',
          title: `Slide ${slide.number} has the same title as slide ${seenTitles.get(key)}`,
          description: 'Unique slide titles help people know where they are. Add a distinguishing word, such as “(continued)” or the topic of the slide.',
          location: loc(slide),
          fix: { kind: 'input', label: 'Change title', inputLabel: 'New slide title', suggestion: `${st} (continued)` },
          target: { slide: slide.number - 1 },
        }));
      } else seenTitles.set(key, slide.number);
      // Title should be read first (1.3.2)
      const content = shs.filter((s) => s.localName !== 'cxnSp');
      if (content.indexOf(titleShape) > 0) {
        issues.push(makeIssue({
          id: `pptx-titleorder-${slide.number}`, rule: 'title-order', wcag: ['1.3.2'],
          title: `Slide ${slide.number}: title is not read first`,
          description: 'Screen readers follow the order in the Selection Pane / Reading Order pane. The title should come first.',
          location: loc(slide), fix: { kind: 'auto', label: 'Move title to the start of the reading order' }, target: { slide: slide.number - 1 },
        }));
      }
    }
    if (shs.length >= 3) readingOrder.push(`Slide ${slide.number}: ` + shs.map((s, k) => `${k + 1}. ${describeShape(s)}`).join(', '));

    // Pictures, charts, graphics (1.1.1)
    shs.forEach((shape, k) => {
      const isPic = shape.localName === 'pic';
      const isGraphic = shape.localName === 'graphicFrame' && !X.first(shape, 'a', 'tbl');
      const grp = shape.localName === 'grpSp';
      if (!isPic && !isGraphic && !grp) return;
      const targets = grp ? X.all(shape, 'p', 'pic') : [shape];
      targets.forEach((pic) => {
        const c = cNvPrOf(pic);
        if (!c || X.isDecorative(c)) return;
        const descr = (c.getAttribute('descr') || '').trim();
        const blip = X.first(pic, 'a', 'blip');
        const rid = blip ? X.attr(blip, 'r', 'embed') : null;
        const rel = rid && slide.relMap[rid];
        const imagePart = rel && !rel.external ? path.normalize(path.join(path.dirname(slide.path), rel.target)) : null;
        const kind = isGraphic ? 'chart or graphic' : 'picture';
        let chartPart = null;
        const chartEl = X.first(pic, 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'chart');
        const crel = chartEl && slide.relMap[X.attr(chartEl, 'r', 'id')];
        if (crel && !crel.external) chartPart = path.normalize(path.join(path.dirname(slide.path), crel.target));
        const caption = D.findCaption(shs.filter((o) => o !== shape).flatMap((o) => shapeText(o).split('\n')));
        const media = X.first(pic, 'a', 'videoFile') || X.first(pic, 'a', 'audioFile') || X.all(pic, '*', 'media').length;
        if (media) {
          issues.push(makeIssue({
            id: `pptx-media-${slide.number}-${c.getAttribute('id')}`, rule: 'media', wcag: ['1.2.2'], severity: 'warning',
            title: 'Video or audio needs captions or a transcript', description: 'Check this media has captions (for video) or a transcript (for audio).',
            location: loc(slide, describeShape(pic)), fix: { kind: 'manual' },
          }));
        }
        const fileLike = /\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(descr) || /^(picture|image|img|graphic|chart)[\s_-]*\d*$/i.test(descr);
        if (!descr || fileLike) {
          issues.push(makeIssue({
            id: `pptx-img-${slide.number}-${c.getAttribute('id')}`, rule: descr ? 'img-alt-quality' : 'img-alt', wcag: ['1.1.1'], severity: descr ? 'warning' : 'error',
            title: descr ? `${cap(kind)} alt text looks like a file name` : `${cap(kind)} has no alt text`,
            description: `Describe what this ${kind} shows or means, or mark it as decorative if it adds no information.`,
            location: loc(slide, describeShape(pic)),
            fix: { kind: imagePart ? 'ai' : 'input', label: 'Add alt text', inputLabel: 'Alt text', allowDecorative: true, multiline: true, ...(caption ? { suggestion: caption, suggestionSource: 'caption' } : {}) },
            target: { slide: slide.number - 1, shapeId: c.getAttribute('id'), imagePart, chartPart, context: `Slide title: ${st}\n${slideText(slide)}` },
          }));
        }
      });
    });

    // Tables (1.3.1)
    X.all(slide.doc, 'a', 'tbl').forEach((tbl, k) => {
      const tblPr = X.kid(tbl, 'a', 'tblPr');
      const rows = X.kids(tbl, 'a', 'tr');
      if (rows.length < 2) return;
      if (!tblPr || tblPr.getAttribute('firstRow') !== '1') {
        issues.push(makeIssue({
          id: `pptx-tbl-${slide.number}-${k}`, rule: 'table-header', wcag: ['1.3.1'], title: 'Table has no header row',
          description: 'Turn on “Header Row” so screen readers announce the column headings for each cell.',
          location: loc(slide, `table ${k + 1}`), fix: { kind: 'auto', label: 'Turn on header row' }, target: { slide: slide.number - 1, tbl: k },
        }));
      }
      if (X.all(tbl, 'a', 'tc').some((tc) => tc.getAttribute('gridSpan') > 1 || tc.getAttribute('rowSpan') > 1 || tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1')) {
        issues.push(makeIssue({
          id: `pptx-tblmerge-${slide.number}-${k}`, rule: 'table-complex', wcag: ['1.3.1'], severity: 'warning', title: 'Table has merged cells',
          description: 'Merged cells are hard to follow with a screen reader. Consider a simpler table.', location: loc(slide, `table ${k + 1}`), fix: { kind: 'manual' },
        }));
      }
    });

    // Links (2.4.4)
    X.all(slide.doc, 'a', 'hlinkClick').forEach((h, k) => {
      const rPr = h.parentNode;
      if (!rPr || rPr.localName !== 'rPr') return;
      const r = rPr.parentNode;
      const t = X.text(X.kid(r, 'a', 't')).trim();
      const rel = slide.relMap[X.attr(h, 'r', 'id')];
      const url = rel ? rel.target : '';
      if (!t || (!LINK_TEXT_BAD.test(t) && !isUrlText(t))) return;
      issues.push(makeIssue({
        id: `pptx-link-${slide.number}-${k}`, rule: 'link-text', wcag: ['2.4.4'], severity: isUrlText(t) ? 'warning' : 'error',
        title: isUrlText(t) ? 'Link text is a raw web address' : `Link text “${t}” is not descriptive`,
        description: 'Link text should say where the link goes so it makes sense on its own.',
        location: { ...loc(slide), highlightText: t },
        fix: { kind: 'ai', label: 'Change link text', inputLabel: 'New link text', suggestion: suggestLinkText(url, t) },
        target: { slide: slide.number - 1, link: k, url, current: t, context: slideText(slide) },
      }));
    });

    // Typed lists (1.3.1)
    for (const shape of shs) {
      const body = X.first(shape, 'p', 'txBody');
      if (!body) continue;
      const ps = X.kids(body, 'a', 'p');
      let group = [];
      const flush = () => {
        if (group.length >= 2) {
          const ordered = /^\s*\(?[0-9a-zA-Z]{1,3}[.)]/.test(paraText(ps[group[0]]));
          issues.push(makeIssue({
            id: `pptx-list-${slide.number}-${cNvPrOf(shape).getAttribute('id')}-${group[0]}`, rule: 'list-manual', wcag: ['1.3.1'],
            title: `List typed by hand (${group.length} items)`, description: 'Lines start with typed bullets or numbers. Fixing turns them into real list formatting.',
            location: loc(slide, describeShape(shape)), fix: { kind: 'auto', label: `Convert to ${ordered ? 'numbered' : 'bulleted'} list` },
            target: { slide: slide.number - 1, shapeId: cNvPrOf(shape).getAttribute('id'), paras: [...group], ordered },
          }));
        }
        group = [];
      };
      ps.forEach((p, i) => {
        const pPr = X.kid(p, 'a', 'pPr');
        const hasBullet = pPr && (X.kid(pPr, 'a', 'buChar') || X.kid(pPr, 'a', 'buAutoNum'));
        const t = paraText(p);
        if (!hasBullet && LIST_MARKER.test(t) && t.replace(LIST_MARKER, '').trim()) group.push(i); else flush();
      });
      flush();
    }

    // Contrast (1.4.3)
    const groups = new Map();
    for (const run of textRuns(ctx, slide)) {
      const need = C.requiredRatio(run.sizePt, run.bold);
      const ratio = C.ratio(run.fg, run.bg);
      if (ratio === null || ratio >= need) continue;
      const key = `${run.fg}|${run.bg}|${need}`;
      if (!groups.has(key)) groups.set(key, { ...run, need, ratio, count: 0 });
      groups.get(key).count++;
    }
    let gi = 0;
    for (const g of groups.values()) {
      const to = C.fixForeground(g.fg, g.bg, g.need);
      issues.push(makeIssue({
        id: `pptx-contrast-${slide.number}-${gi++}`, rule: 'contrast', wcag: ['1.4.3'],
        title: `Low colour contrast (${g.ratio.toFixed(2)}:1, needs ${g.need}:1)`,
        description: `Text colour #${g.fg} on #${g.bg} is hard to read${g.count > 1 ? ` (${g.count} places on this slide)` : ''}. Fixing changes the text colour to #${to}.`,
        location: { ...loc(slide, describeShape(g.shape)), highlightText: g.text.trim().slice(0, 80) },
        fix: { kind: 'auto', label: `Change colour to #${to}` },
        target: { slide: slide.number - 1, fg: g.fg, bg: g.bg, need: g.need, to, swatch: { fg: g.fg, bg: g.bg, to } },
      }));
    }
  }

  // Charts: describe them from their own data (no AI needed).
  for (const iss of issues) {
    const cp = iss.target && iss.target.chartPart;
    if (!cp || iss.fix.suggestionSource || !ctx.zip.file(cp)) continue;
    const d = D.describeChart(await ctx.zip.file(cp).async('string'));
    if (d) { iss.fix.suggestion = d; iss.fix.suggestionSource = 'chart data'; }
  }
  for (const iss of issues) {
    if (iss.rule === 'slide-title' && iss.fix.suggestion) iss.fix.suggestionSource = 'first line of the slide';
    if (iss.rule === 'link-text' && iss.fix.suggestion) iss.fix.suggestionSource = 'web address';
    if (iss.rule === 'slide-title-dup') iss.fix.suggestionSource = 'existing title';
  }

  return {
    format: 'pptx',
    meta: { title: title || null, language: presLang || null, slides: ctx.slides.length },
    humanCheckDetails: readingOrder.length ? { 'hc-reading-order': readingOrder } : {},
    issues,
  };
}
function paraText(p) {
  return X.all(p, 'a', 't').map((t) => t.textContent).join('');
}
function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// ---------------------------------------------------------------------------
async function fix(buffer, issue, input = {}, opts = {}) {
  const ctx = await load(buffer);
  const t = issue.target || {};
  const slide = t.slide !== undefined ? ctx.slides[t.slide] : null;
  const value = (input.value !== undefined ? input.value : issue.fix.suggestion || '').trim();
  switch (issue.rule) {
    case 'doc-title': {
      if (!value) throw new Error('Enter a title.');
      if (!ctx.core) {
        ctx.core = X.parse(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${X.NS.cp}" xmlns:dc="${X.NS.dc}" xmlns:dcterms="${X.NS.dcterms}" xmlns:xsi="${X.NS.xsi}"></cp:coreProperties>`);
        ctx.newCore = true;
      }
      let el = X.first(ctx.core, 'dc', 'title');
      if (!el) { el = ctx.core.createElementNS(X.NS.dc, 'dc:title'); ctx.core.documentElement.appendChild(el); }
      el.textContent = value;
      ctx.dirty.add('core');
      break;
    }
    case 'doc-lang': {
      const lang = value || opts.defaultLanguage || 'en-GB';
      if (!/^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$/.test(lang)) throw new Error('Use a language code such as en-GB.');
      let dts = X.first(ctx.pres, 'p', 'defaultTextStyle');
      if (!dts) {
        dts = X.el(ctx.pres, 'p', 'defaultTextStyle');
        const after = X.first(ctx.pres, 'p', 'notesSz');
        ctx.pres.documentElement.insertBefore(dts, after ? after.nextSibling : null);
      }
      let defPPr = X.kid(dts, 'a', 'defPPr');
      if (!defPPr) { defPPr = X.el(ctx.pres, 'a', 'defPPr'); dts.insertBefore(defPPr, dts.firstChild); }
      let d = X.kid(defPPr, 'a', 'defRPr');
      if (!d) { d = X.el(ctx.pres, 'a', 'defRPr'); defPPr.appendChild(d); }
      d.setAttribute('lang', lang);
      for (const e of X.all(dts, 'a', 'defRPr')) e.setAttribute('lang', lang);
      ctx.dirty.add('pres');
      for (const s of ctx.slides) {
        for (const r of X.all(s.doc, 'a', 'r')) {
          let rp = X.kid(r, 'a', 'rPr');
          if (!rp) { rp = X.el(s.doc, 'a', 'rPr'); r.insertBefore(rp, r.firstChild); }
          if (!rp.getAttribute('lang')) rp.setAttribute('lang', lang);
        }
        for (const e of X.all(s.doc, 'a', 'endParaRPr')) if (!e.getAttribute('lang')) e.setAttribute('lang', lang);
        s.dirty = true;
      }
      break;
    }
    case 'slide-title':
    case 'slide-title-dup': {
      if (!value) throw new Error('Enter a slide title.');
      const existing = shapes(slide).find(isTitleShape);
      if (existing) setShapeText(slide.doc, existing, value);
      else addHiddenTitle(ctx, slide, value);
      slide.dirty = true;
      break;
    }
    case 'title-order': {
      const tree = X.first(slide.doc, 'p', 'spTree');
      const ts = shapes(slide).find(isTitleShape);
      const firstShape = shapes(slide)[0];
      tree.insertBefore(ts, firstShape);
      slide.dirty = true;
      break;
    }
    case 'img-alt':
    case 'img-alt-quality': {
      const c = X.all(slide.doc, 'p', 'cNvPr').find((e) => e.getAttribute('id') === t.shapeId);
      if (!c) throw new Error('Picture not found. Re-check and try again.');
      if (input.decorative) X.markDecorative(c);
      else { if (!value) throw new Error('Enter alt text, or mark it as decorative.'); c.setAttribute('descr', value); }
      slide.dirty = true;
      break;
    }
    case 'table-header': {
      const tbl = X.all(slide.doc, 'a', 'tbl')[t.tbl];
      let tblPr = X.kid(tbl, 'a', 'tblPr');
      if (!tblPr) { tblPr = X.el(slide.doc, 'a', 'tblPr'); tbl.insertBefore(tblPr, tbl.firstChild); }
      tblPr.setAttribute('firstRow', '1');
      slide.dirty = true;
      break;
    }
    case 'link-text': {
      const h = X.all(slide.doc, 'a', 'hlinkClick')[t.link];
      if (!value) throw new Error('Enter the new link text.');
      const r = h.parentNode.parentNode;
      X.kid(r, 'a', 't').textContent = value;
      slide.dirty = true;
      break;
    }
    case 'list-manual': {
      const shape = shapes(slide).find((s) => cNvPrOf(s) && cNvPrOf(s).getAttribute('id') === t.shapeId);
      const ps = X.kids(X.first(shape, 'p', 'txBody'), 'a', 'p');
      for (const i of t.paras) {
        const p = ps[i];
        stripMarker(p);
        let pPr = X.kid(p, 'a', 'pPr');
        if (!pPr) { pPr = X.el(slide.doc, 'a', 'pPr'); p.insertBefore(pPr, p.firstChild); }
        for (const k of ['buNone', 'buChar', 'buAutoNum', 'buFont']) { const e = X.kid(pPr, 'a', k); if (e) X.remove(e); }
        pPr.setAttribute('marL', '342900'); pPr.setAttribute('indent', '-342900');
        // Bullet elements must come after spacing / colour elements and before tabLst/defRPr.
        const before = X.kid(pPr, 'a', 'tabLst') || X.kid(pPr, 'a', 'defRPr') || X.kid(pPr, 'a', 'extLst');
        if (t.ordered) pPr.insertBefore(X.el(slide.doc, 'a', 'buAutoNum', { type: 'arabicPeriod' }), before);
        else { pPr.insertBefore(X.el(slide.doc, 'a', 'buFont', { typeface: 'Arial' }), before); pPr.insertBefore(X.el(slide.doc, 'a', 'buChar', { char: '•' }), before); }
      }
      slide.dirty = true;
      break;
    }
    case 'contrast': {
      for (const run of textRuns(ctx, slide)) {
        if (run.fg !== t.fg || run.bg !== t.bg || C.requiredRatio(run.sizePt, run.bold) !== t.need) continue;
        const rPr = X.kid(run.r, 'a', 'rPr');
        const solid = X.kid(rPr, 'a', 'solidFill');
        while (solid.firstChild) solid.removeChild(solid.firstChild);
        solid.appendChild(X.el(slide.doc, 'a', 'srgbClr', { val: t.to }));
      }
      slide.dirty = true;
      break;
    }
    default:
      throw new Error('This issue cannot be fixed automatically.');
  }
  return save(ctx);
}
function stripMarker(p) {
  const m = paraText(p).match(LIST_MARKER);
  if (!m) return;
  let remaining = m[0].length;
  for (const tEl of X.all(p, 'a', 't')) {
    if (remaining <= 0) break;
    const s = tEl.textContent;
    if (s.length <= remaining) { remaining -= s.length; tEl.textContent = ''; } else { tEl.textContent = s.slice(remaining); remaining = 0; }
  }
}
function setShapeText(doc, shape, value) {
  let body = X.first(shape, 'p', 'txBody');
  if (!body) {
    body = X.el(doc, 'p', 'txBody');
    body.appendChild(X.el(doc, 'a', 'bodyPr')); body.appendChild(X.el(doc, 'a', 'lstStyle'));
    shape.appendChild(body);
  }
  for (const p of X.kids(body, 'a', 'p')) X.remove(p);
  const p = X.el(doc, 'a', 'p');
  const r = X.el(doc, 'a', 'r');
  const lang = X.first(doc, 'a', 'rPr');
  const rPr = X.el(doc, 'a', 'rPr', lang && lang.getAttribute('lang') ? { lang: lang.getAttribute('lang') } : {});
  const t = X.el(doc, 'a', 't'); t.textContent = value;
  r.appendChild(rPr); r.appendChild(t); p.appendChild(r); body.appendChild(p);
}
function addHiddenTitle(ctx, slide, value) {
  const doc = slide.doc;
  const tree = X.first(doc, 'p', 'spTree');
  const ids = X.all(doc, 'p', 'cNvPr').map((c) => +c.getAttribute('id') || 0);
  const id = Math.max(1, ...ids) + 1;
  const sldSz = X.first(ctx.pres, 'p', 'sldSz');
  const cx = sldSz ? +sldSz.getAttribute('cx') : 9144000;
  const sp = X.el(doc, 'p', 'sp');
  const nv = X.el(doc, 'p', 'nvSpPr');
  nv.appendChild(X.el(doc, 'p', 'cNvPr', { id: String(id), name: `Title ${id}` }));
  const cnv = X.el(doc, 'p', 'cNvSpPr'); cnv.appendChild(X.el(doc, 'a', 'spLocks', { noGrp: '1' })); nv.appendChild(cnv);
  const nvPr = X.el(doc, 'p', 'nvPr'); nvPr.appendChild(X.el(doc, 'p', 'ph', { type: 'title' })); nv.appendChild(nvPr);
  sp.appendChild(nv);
  const spPr = X.el(doc, 'p', 'spPr');
  const xfrm = X.el(doc, 'a', 'xfrm');
  xfrm.appendChild(X.el(doc, 'a', 'off', { x: '457200', y: '-1028700' }));
  xfrm.appendChild(X.el(doc, 'a', 'ext', { cx: String(Math.max(914400, cx - 914400)), cy: '685800' }));
  spPr.appendChild(xfrm);
  sp.appendChild(spPr);
  setShapeText(doc, sp, value);
  const firstShape = shapes(slide)[0];
  tree.insertBefore(sp, firstShape || null);
}

async function getImage(buffer, target) {
  if (!target || !target.imagePart) return null;
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file(target.imagePart);
  if (!f) return null;
  return { data: await f.async('nodebuffer'), mime: mimeOf(target.imagePart) };
}

module.exports = { audit, fix, getImage };
