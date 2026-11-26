'use strict';
// PDF accessibility audit (WCAG 2.1 AA + key PDF/UA checks) and the fixes that can
// be made safely without the source document.
//
// Honest limits: this module does NOT invent a tag tree for an untagged PDF. Auto-
// generated tags are usually wrong (reading order, tables, lists) and would make a
// document *look* compliant to checkers while still failing screen reader users.
// Untagged PDFs are flagged with advice to fix the source file (which this app can
// do for Word and PowerPoint) and export again with tags.
const path = require('path');
const { pathToFileURL } = require('url');
const { PDFDocument, PDFName, PDFDict, PDFArray, PDFNumber, PDFHexString, PDFString, PDFRef, PDFBool } = require('pdf-lib');
const { makeIssue } = require('./issues');

const N = (s) => PDFName.of(s);
const HEADING = /^H([1-6])$/;

let pdfjsPromise = null;
function pdfjs() {
  if (!pdfjsPromise) {
    const file = require.resolve('pdfjs-dist/legacy/build/pdf.mjs');
    pdfjsPromise = import(pathToFileURL(file).href);
  }
  return pdfjsPromise;
}
function standardFontDir() {
  return path.join(path.dirname(require.resolve('pdfjs-dist/package.json')), 'standard_fonts') + path.sep;
}

// ---------------------------------------------------------------------------
async function loadDoc(buffer, forEdit) {
  try {
    return await PDFDocument.load(buffer, { updateMetadata: !!forEdit, ignoreEncryption: !forEdit, throwOnInvalidObject: false });
  } catch (e) {
    if (/encrypt/i.test(e.message)) throw new Error('This PDF is password protected or encrypted, so it cannot be changed. Remove the protection in the program that created it, then try again.');
    throw new Error('This PDF could not be read: ' + e.message);
  }
}
function str(obj) {
  if (!obj) return '';
  if (obj instanceof PDFHexString || obj instanceof PDFString) return obj.decodeText();
  return '';
}
function lookupDict(ctx, obj) {
  const v = obj instanceof PDFRef ? ctx.lookup(obj) : obj;
  return v instanceof PDFDict ? v : null;
}
function asArray(ctx, obj) {
  const v = obj instanceof PDFRef ? ctx.lookup(obj) : obj;
  if (v instanceof PDFArray) return v.asArray();
  return v === undefined || v === null ? [] : [obj];
}

/** Walk the structure tree, yielding { el, type, parentPg, depth }. */
function walkStruct(pdf) {
  const ctx = pdf.context;
  const root = pdf.catalog.lookup(N('StructTreeRoot'));
  if (!(root instanceof PDFDict)) return null;
  const roleMap = root.lookup(N('RoleMap'));
  const mapRole = (s) => {
    let cur = s, guard = 0;
    while (roleMap instanceof PDFDict && guard++ < 10) {
      const m = roleMap.get(N(cur));
      if (!(m instanceof PDFName)) break;
      cur = m.decodeText();
    }
    return cur;
  };
  const out = [];
  const visit = (obj, pg, depth, parent) => {
    for (const k of asArray(ctx, obj)) {
      const d = lookupDict(ctx, k);
      if (!d) continue;
      const type = d.get(N('Type'));
      if (type instanceof PDFName && (type.decodeText() === 'MCR' || type.decodeText() === 'OBJR')) continue;
      const s = d.get(N('S'));
      if (!(s instanceof PDFName)) continue;
      const pgRef = d.get(N('Pg')) || pg;
      const rec = { el: d, raw: s.decodeText(), type: mapRole(s.decodeText()), pg: pgRef, depth, parent, kids: [] };
      out.push(rec);
      if (parent) parent.kids.push(rec);
      visit(d.get(N('K')), pgRef, depth + 1, rec);
    }
  };
  visit(root.get(N('K')), null, 0, null);
  return out;
}
/** MCIDs (with their page refs) under a struct element. */
function mcidsOf(pdf, rec) {
  const ctx = pdf.context;
  const out = [];
  const visit = (obj, pg) => {
    for (const k of asArray(ctx, obj)) {
      const v = k instanceof PDFRef ? ctx.lookup(k) : k;
      if (v instanceof PDFNumber) out.push({ mcid: v.asNumber(), pg });
      else if (v instanceof PDFDict) {
        const t = v.get(N('Type'));
        if (t instanceof PDFName && t.decodeText() === 'MCR') {
          const m = v.get(N('MCID')); if (m instanceof PDFNumber) out.push({ mcid: m.asNumber(), pg: v.get(N('Pg')) || pg });
        } else if (v.get(N('S'))) visit(v.get(N('K')), v.get(N('Pg')) || pg);
      }
    }
  };
  visit(rec.el.get(N('K')), rec.pg);
  return out;
}

/** Extract page text and marked-content text with pdf.js (read-only). */
async function extractText(buffer) {
  const lib = await pdfjs();
  const task = lib.getDocument({ data: new Uint8Array(buffer), standardFontDataUrl: standardFontDir(), isEvalSupported: false, verbosity: 0 });
  const doc = await task.promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const tc = await page.getTextContent({ includeMarkedContent: true });
    const mc = {};
    const stack = [];
    let text = '';
    for (const it of tc.items) {
      if (it.type === 'beginMarkedContentProps' || it.type === 'beginMarkedContent') {
        const m = it.id ? String(it.id).match(/mc(\d+)$/) : null;
        stack.push(m ? +m[1] : null);
      } else if (it.type === 'endMarkedContent') stack.pop();
      else if (typeof it.str === 'string') {
        text += it.str + (it.hasEOL ? '\n' : '');
        for (const id of stack) if (id !== null) mc[id] = (mc[id] || '') + it.str + (it.hasEOL ? ' ' : '');
      }
    }
    const ops = await page.getOperatorList();
    const imageOps = new Set([lib.OPS.paintImageXObject, lib.OPS.paintInlineImageXObject, lib.OPS.paintImageMaskXObject].filter(Boolean));
    const images = ops.fnArray.filter((f) => imageOps.has(f)).length;
    pages.push({ text, mc, images });
    page.cleanup();
  }
  await task.destroy();
  return pages;
}

function fontsNotEmbedded(pdf) {
  const ctx = pdf.context;
  const missing = new Set();
  const checkFont = (font) => {
    const sub = font.get(N('Subtype'));
    const subtype = sub instanceof PDFName ? sub.decodeText() : '';
    if (subtype === 'Type3') return;
    let desc = font.lookup(N('FontDescriptor'));
    if (subtype === 'Type0') {
      const df = font.lookup(N('DescendantFonts'));
      const d0 = df instanceof PDFArray ? lookupDict(ctx, df.get(0)) : null;
      desc = d0 ? d0.lookup(N('FontDescriptor')) : null;
    }
    const base = font.get(N('BaseFont'));
    const name = base instanceof PDFName ? base.decodeText() : 'Unknown font';
    if (!(desc instanceof PDFDict) || !(desc.get(N('FontFile')) || desc.get(N('FontFile2')) || desc.get(N('FontFile3')))) missing.add(name.replace(/^[A-Z]{6}\+/, ''));
  };
  for (const page of pdf.getPages()) {
    const res = page.node.Resources();
    const fonts = res ? res.lookup(N('Font')) : null;
    if (fonts instanceof PDFDict) for (const [, ref] of fonts.entries()) { const f = lookupDict(ctx, ref); if (f) checkFont(f); }
  }
  return [...missing];
}

// ---------------------------------------------------------------------------
async function audit(buffer, opts = {}) {
  const pdf = await loadDoc(buffer, false);
  const issues = [];
  const pages = pdf.getPages();
  const pageRefs = pages.map((p) => p.ref.toString());
  const pageOf = (ref) => (ref ? pageRefs.indexOf(ref.toString()) + 1 : 0);
  const fileTitle = (opts.fileName || 'Document').replace(/\.[^.]+$/, '');

  if (pdf.isEncrypted) {
    issues.push(makeIssue({
      id: 'pdf-encrypted', rule: 'encrypted', wcag: ['4.1.2'], pdfua: '7.16',
      title: 'PDF is encrypted', description: 'Encryption can stop screen readers reading the text and stops this app from making fixes. Save an unprotected copy from the program that created it.',
      fix: { kind: 'manual' },
    }));
  }

  let text = [];
  try { text = await extractText(buffer); } catch (e) { if (process.env.ADC_DEBUG) console.error(e); text = pages.map(() => ({ text: '', mc: {}, images: 0, failed: true })); }

  // Title + DisplayDocTitle (2.4.2, PDF/UA 7.1)
  const title = (pdf.getTitle() || '').trim();
  const vp = pdf.catalog.lookup(N('ViewerPreferences'));
  const displayTitle = vp instanceof PDFDict && vp.get(N('DisplayDocTitle')) === PDFBool.True;
  const firstLine = (text[0] && text[0].text.split('\n').map((s) => s.trim()).find((s) => s.length > 3)) || '';
  if (!title) {
    issues.push(makeIssue({
      id: 'pdf-title', rule: 'doc-title', wcag: ['2.4.2'], pdfua: '7.1',
      title: 'Document title is missing', description: 'Screen readers and the window title bar show the document title. Without it people hear the file name.',
      fix: { kind: 'auto', label: 'Set title', suggestion: firstLine.slice(0, 120) || fileTitle.replace(/[_-]+/g, ' '), editable: true, inputLabel: 'Document title' },
    }));
  } else if (!displayTitle) {
    issues.push(makeIssue({
      id: 'pdf-displaytitle', rule: 'display-title', wcag: ['2.4.2'], pdfua: '7.1',
      title: 'Window shows the file name instead of the title', description: 'The “Display document title” setting is off, so the title bar shows the file name.',
      fix: { kind: 'auto', label: 'Show title in window' },
    }));
  }

  // Language (3.1.1)
  const lang = str(pdf.catalog.lookup(N('Lang')));
  if (!lang) {
    issues.push(makeIssue({
      id: 'pdf-lang', rule: 'doc-lang', wcag: ['3.1.1'], pdfua: '7.2',
      title: 'Document language is not set', description: 'Without a language, screen readers may read the text with the wrong pronunciation.',
      fix: { kind: 'auto', label: 'Set language', suggestion: opts.defaultLanguage || 'en-GB', editable: true, inputLabel: 'Language code (for example en-GB)' },
    }));
  }

  // Scanned pages (1.1.1 / 1.4.5)
  const noText = text.map((p, i) => (!p.failed && !p.text.trim() && p.images > 0 ? i + 1 : 0)).filter(Boolean);
  if (noText.length) {
    issues.push(makeIssue({
      id: 'pdf-scanned', rule: 'scanned', wcag: ['1.1.1', '1.4.5'],
      title: `${noText.length === pages.length ? 'Whole document is' : `${noText.length} page${noText.length > 1 ? 's are' : ' is'}`} an image with no real text`,
      description: 'Scanned pages are pictures, so screen readers find nothing to read. Run OCR (text recognition) in Acrobat or similar, or recreate the document from its source, then check it again.',
      location: { label: `Page${noText.length > 1 ? 's' : ''} ${listPages(noText)}` }, fix: { kind: 'manual' },
    }));
  }

  // Tags (1.3.1, PDF/UA 7.1)
  const struct = walkStruct(pdf);
  const markInfo = pdf.catalog.lookup(N('MarkInfo'));
  const marked = markInfo instanceof PDFDict && markInfo.get(N('Marked')) === PDFBool.True;
  const headings = [];
  if (!struct) {
    issues.push(makeIssue({
      id: 'pdf-untagged', rule: 'untagged', wcag: ['1.3.1', '1.3.2', '4.1.2'], pdfua: '7.1',
      title: 'PDF is not tagged',
      description: 'Tags tell screen readers what is a heading, list, table or image, and the reading order. They cannot be reliably added automatically. Best fix: correct the original Word or PowerPoint file (this app can do that), then save as PDF with “Document structure tags for accessibility” turned on. Otherwise add tags in Adobe Acrobat Pro (Accessibility > Autotag) and review them.',
      fix: { kind: 'manual' },
    }));
  } else {
    if (!marked) {
      issues.push(makeIssue({
        id: 'pdf-marked', rule: 'mark-info', wcag: ['1.3.1'], pdfua: '7.1',
        title: 'PDF is not flagged as tagged', description: 'The document has tags, but the “Marked” flag is off, so some assistive technology ignores them.',
        fix: { kind: 'auto', label: 'Set tagged flag' },
      }));
    }
    // Figures (1.1.1)
    const figsPerPage = {};
    struct.filter((r) => r.type === 'Figure').forEach((r, i) => {
      const alt = str(r.el.lookup(N('Alt'))) || str(r.el.lookup(N('ActualText')));
      const pg = pageOf(r.pg);
      const nth = figsPerPage[pg] = (figsPerPage[pg] || 0) + 1;
      if (alt.trim()) return;
      // Offline suggestion: the matching "Figure 1: ..." caption on the page.
      const caps = pg ? [...text[pg - 1].text.matchAll(/(?:^|[\s.])(?:Figure|Fig\.|Chart|Graph|Image|Photo|Diagram|Map)\s*\d+\s*[:.\-–—]\s*([^\n]{3,200}?)(?=\s*(?:$|\n|(?:Figure|Fig\.|Chart|Graph|Image|Photo|Diagram|Map)\s*\d+\s*[:.\-–—])|\.\s)/g)].map((m) => m[1].trim()) : [];
      const caption = caps[nth - 1] || '';
      issues.push(makeIssue({
        id: `pdf-fig-${i}`, rule: 'img-alt', wcag: ['1.1.1'], pdfua: '7.3',
        title: 'Figure has no alt text', description: 'Add a short description of what this figure shows or means. If it is purely decorative, it should be an artifact instead (fix that in the source or in Acrobat).',
        location: { label: pg ? `Page ${pg} · figure ${i + 1}` : `Figure ${i + 1}`, page: pg, excerpt: pg ? short(text[pg - 1].text, 200) : '' },
        fix: { kind: pg ? 'ai' : 'input', label: 'Add alt text', inputLabel: 'Alt text', multiline: true, ...(caption ? { suggestion: caption, suggestionSource: 'caption' } : {}) },
        target: { figure: i, page: pg ? pg - 1 : null, context: pg ? short(text[pg - 1].text, 800) : '' },
      }));
    });
    // Headings (1.3.1, 2.4.6)
    let prev = 0;
    struct.forEach((r, idx) => {
      const m = r.type.match(HEADING);
      if (!m) return;
      const level = +m[1];
      const pg = pageOf(r.pg);
      const t = short(mcidsOf(pdf, r).map(({ mcid, pg: p }) => { const n = pageOf(p || r.pg); return n && text[n - 1] ? text[n - 1].mc[mcid] || '' : ''; }).join(' '), 120) || str(r.el.lookup(N('T')));
      headings.push({ level, title: t, page: pg, ref: r.pg });
      if (prev && level > prev + 1) {
        issues.push(makeIssue({
          id: `pdf-hskip-${idx}`, rule: 'heading-skip', wcag: ['1.3.1', '2.4.6'], pdfua: '7.4.2',
          title: `Heading level skipped (H${prev} to H${level})`, description: `Heading levels should go down one step at a time. This tag will be changed to H${prev + 1}.`,
          location: { label: pg ? `Page ${pg}` : 'Heading', page: pg, excerpt: t, highlightText: t }, fix: { kind: 'auto', label: `Change to H${prev + 1}` },
          target: { structIndex: idx, level: prev + 1 },
        }));
        prev = prev + 1;
      } else prev = level;
    });
    if (!headings.length && pages.length >= 3 && !struct.some((r) => r.type === 'H')) {
      issues.push(makeIssue({
        id: 'pdf-noheadings', rule: 'headings-none', wcag: ['1.3.1', '2.4.6'], severity: 'warning',
        title: 'No headings are tagged', description: 'Headings help people navigate. Apply heading styles in the source document and export again, or retag headings in Acrobat.', fix: { kind: 'manual' },
      }));
    }
    // Tables (1.3.1)
    struct.filter((r) => r.type === 'Table').forEach((t, i) => {
      const cells = [];
      const collect = (rec) => { for (const k of rec.kids) { if (k.type === 'TH' || k.type === 'TD') cells.push(k); else collect(k); } };
      collect(t);
      if (!cells.length || cells.some((c) => c.type === 'TH')) return;
      const pg = pageOf(t.pg);
      issues.push(makeIssue({
        id: `pdf-tbl-${i}`, rule: 'table-header', wcag: ['1.3.1'], pdfua: '7.5',
        title: 'Table has no header cells', description: 'No cells are tagged as headers (TH). Fixing tags the first row as column headers. Check this is right for the table.',
        location: { label: pg ? `Page ${pg} · table ${i + 1}` : `Table ${i + 1}`, page: pg }, fix: { kind: 'auto', label: 'Tag first row as headers' }, target: { table: i },
      }));
    });
  }

  // Links and annotations (2.4.4, 2.4.3, PDF/UA 7.18)
  let linksNoContents = 0, pagesNoTabs = 0;
  pages.forEach((page) => {
    const annots = page.node.Annots();
    if (!annots || !annots.size()) return;
    const tabs = page.node.get(N('Tabs'));
    if (!(tabs instanceof PDFName) || tabs.decodeText() !== 'S') pagesNoTabs++;
    for (const a of annots.asArray()) {
      const d = lookupDict(pdf.context, a);
      if (!d) continue;
      const st = d.get(N('Subtype'));
      if (st instanceof PDFName && st.decodeText() === 'Link' && !str(d.lookup(N('Contents'))).trim()) linksNoContents++;
    }
  });
  if (linksNoContents) {
    issues.push(makeIssue({
      id: 'pdf-linkcontents', rule: 'link-contents', wcag: ['2.4.4', '4.1.2'], pdfua: '7.18.5', severity: 'warning',
      title: `${linksNoContents} link${linksNoContents > 1 ? 's have' : ' has'} no description`, description: 'Link annotations should have a description (Contents) that assistive technology can announce. Fixing adds the link address or target page.',
      fix: { kind: 'auto', label: 'Add link descriptions' },
    }));
  }
  if (pagesNoTabs) {
    issues.push(makeIssue({
      id: 'pdf-tabs', rule: 'tab-order', wcag: ['2.4.3'], pdfua: '7.18.3',
      title: `Tab order is not set on ${pagesNoTabs} page${pagesNoTabs > 1 ? 's' : ''}`, description: 'Pages with links or form fields should use the document structure for tab order, so keyboard users move through them in reading order.',
      fix: { kind: 'auto', label: 'Use document structure for tab order' },
    }));
  }

  // Bookmarks (2.4.5) for long documents
  const outlines = pdf.catalog.lookup(N('Outlines'));
  const hasBookmarks = outlines instanceof PDFDict && outlines.get(N('First'));
  if (pages.length > 20 && !hasBookmarks) {
    const usable = headings.filter((h) => h.page && h.title);
    issues.push(makeIssue({
      id: 'pdf-bookmarks', rule: 'bookmarks', wcag: ['2.4.5'], severity: 'warning',
      title: 'Long document has no bookmarks', description: usable.length ? `Bookmarks help people move around a ${pages.length}-page document. Fixing creates bookmarks from the ${usable.length} tagged headings.` : `Bookmarks help people move around a ${pages.length}-page document. There are no tagged headings to build them from, so add headings in the source and export with bookmarks.`,
      fix: usable.length ? { kind: 'auto', label: 'Create bookmarks from headings' } : { kind: 'manual' },
    }));
  }

  // Fonts (PDF/UA 7.21.4)
  const missingFonts = fontsNotEmbedded(pdf);
  if (missingFonts.length) {
    issues.push(makeIssue({
      id: 'pdf-fonts', rule: 'fonts', wcag: ['1.3.1'], pdfua: '7.21.4.1', severity: 'warning',
      title: `${missingFonts.length} font${missingFonts.length > 1 ? 's are' : ' is'} not embedded`, description: `Not embedded: ${missingFonts.slice(0, 6).join(', ')}${missingFonts.length > 6 ? '…' : ''}. PDF/UA requires embedded fonts so text displays and is extracted reliably. Export again with “embed fonts” on.`,
      fix: { kind: 'manual' },
    }));
  }

  return {
    format: 'pdf',
    meta: { title: title || null, language: lang || null, pages: pages.length, tagged: !!struct, headings: headings.length, figures: struct ? struct.filter((r) => r.type === 'Figure').length : 0 },
    issues,
  };
}
function short(s, n = 140) { s = (s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
function listPages(nums) { return nums.length > 8 ? `${nums.slice(0, 8).join(', ')}…` : nums.join(', '); }

// ---------------------------------------------------------------------------
async function fix(buffer, issue, input = {}, opts = {}) {
  const pdf = await loadDoc(buffer, true);
  const ctx = pdf.context;
  const t = issue.target || {};
  const value = (input.value !== undefined ? input.value : issue.fix.suggestion || '').trim();
  switch (issue.rule) {
    case 'doc-title':
      if (!value) throw new Error('Enter a title.');
      pdf.setTitle(value, { showInWindowTitleBar: true });
      writeXmp(pdf, value);
      break;
    case 'display-title':
      pdf.setTitle(pdf.getTitle(), { showInWindowTitleBar: true });
      writeXmp(pdf, pdf.getTitle());
      break;
    case 'doc-lang': {
      const lang = value || opts.defaultLanguage || 'en-GB';
      if (!/^[a-zA-Z]{2,3}(-[a-zA-Z0-9]{2,8})*$/.test(lang)) throw new Error('Use a language code such as en-GB.');
      pdf.setLanguage(lang);
      break;
    }
    case 'mark-info': {
      let mi = pdf.catalog.lookup(N('MarkInfo'));
      if (!(mi instanceof PDFDict)) { mi = ctx.obj({}); pdf.catalog.set(N('MarkInfo'), mi); }
      mi.set(N('Marked'), PDFBool.True);
      break;
    }
    case 'img-alt': {
      const figs = walkStruct(pdf).filter((r) => r.type === 'Figure');
      const f = figs[t.figure];
      if (!f) throw new Error('Figure not found. Re-check and try again.');
      if (!value) throw new Error('Enter alt text.');
      f.el.set(N('Alt'), PDFHexString.fromText(value));
      break;
    }
    case 'heading-skip': {
      const r = walkStruct(pdf)[t.structIndex];
      r.el.set(N('S'), N(`H${t.level}`));
      break;
    }
    case 'table-header': {
      const tbl = walkStruct(pdf).filter((r) => r.type === 'Table')[t.table];
      const firstRow = findFirst(tbl, (r) => r.type === 'TR');
      if (!firstRow) throw new Error('This table has no rows tagged, so headers cannot be set.');
      for (const c of firstRow.kids) if (c.type === 'TD') { c.el.set(N('S'), N('TH')); setScope(ctx, c.el, 'Column'); }
      break;
    }
    case 'link-contents': {
      const pages = pdf.getPages();
      const refs = pages.map((p) => p.ref.toString());
      for (const page of pages) {
        const annots = page.node.Annots();
        if (!annots) continue;
        for (const a of annots.asArray()) {
          const d = lookupDict(ctx, a);
          const st = d && d.get(N('Subtype'));
          if (!(st instanceof PDFName) || st.decodeText() !== 'Link' || str(d.lookup(N('Contents'))).trim()) continue;
          const action = d.lookup(N('A'));
          let desc = 'Link';
          if (action instanceof PDFDict) {
            const uri = action.lookup(N('URI'));
            if (uri) desc = `Link to ${str(uri) || uri.toString().replace(/^\(|\)$/g, '')}`;
            const dest = action.lookup(N('D'));
            if (dest instanceof PDFArray) { const n = refs.indexOf(dest.get(0).toString()); if (n >= 0) desc = `Link to page ${n + 1}`; }
          }
          const dest = d.lookup(N('Dest'));
          if (dest instanceof PDFArray) { const n = refs.indexOf(dest.get(0).toString()); if (n >= 0) desc = `Link to page ${n + 1}`; }
          d.set(N('Contents'), PDFHexString.fromText(desc));
        }
      }
      break;
    }
    case 'tab-order':
      for (const page of pdf.getPages()) if (page.node.Annots()) page.node.set(N('Tabs'), N('S'));
      break;
    case 'bookmarks': {
      const text = await extractText(buffer);
      const struct = walkStruct(pdf);
      const pages = pdf.getPages();
      const refs = pages.map((p) => p.ref.toString());
      const hs = [];
      for (const r of struct) {
        const m = r.type.match(HEADING);
        if (!m || !r.pg) continue;
        const pgIdx = refs.indexOf(r.pg.toString());
        const title = short(mcidsOf(pdf, r).map(({ mcid, pg }) => { const n = refs.indexOf((pg || r.pg).toString()); return n >= 0 ? text[n].mc[mcid] || '' : ''; }).join(' '), 120) || str(r.el.lookup(N('T')));
        if (pgIdx >= 0 && title) hs.push({ level: +m[1], title, page: pages[pgIdx] });
      }
      buildOutlines(pdf, hs);
      break;
    }
    default:
      throw new Error('This issue cannot be fixed automatically.');
  }
  return Buffer.from(await pdf.save({ useObjectStreams: false }));
}
function findFirst(rec, pred) {
  for (const k of rec.kids) { if (pred(k)) return k; const f = findFirst(k, pred); if (f) return f; }
  return null;
}
function setScope(ctx, el, scope) {
  let a = el.lookup(N('A'));
  if (a instanceof PDFArray) { a = a.lookup(0); }
  if (!(a instanceof PDFDict)) { a = ctx.obj({ O: 'Table' }); el.set(N('A'), a); }
  a.set(N('O'), N('Table')); a.set(N('Scope'), N(scope));
}
function buildOutlines(pdf, headings) {
  const ctx = pdf.context;
  const root = ctx.obj({ Type: 'Outlines' });
  const rootRef = ctx.register(root);
  const stack = [{ level: 0, dict: root, ref: rootRef, kids: [] }];
  const all = [];
  for (const h of headings) {
    while (stack.length > 1 && stack[stack.length - 1].level >= h.level) stack.pop();
    const parent = stack[stack.length - 1];
    const dict = ctx.obj({ Title: PDFHexString.fromText(h.title), Parent: parent.ref, Dest: [h.page.ref, 'XYZ', null, null, null] });
    const node = { level: h.level, dict, ref: ctx.register(dict), kids: [] };
    parent.kids.push(node);
    all.push(node);
    stack.push(node);
  }
  const link = (node) => {
    if (!node.kids.length) return 0;
    node.dict.set(N('First'), node.kids[0].ref);
    node.dict.set(N('Last'), node.kids[node.kids.length - 1].ref);
    node.kids.forEach((k, i) => {
      if (i > 0) k.dict.set(N('Prev'), node.kids[i - 1].ref);
      if (i < node.kids.length - 1) k.dict.set(N('Next'), node.kids[i + 1].ref);
    });
    let count = node.kids.length;
    for (const k of node.kids) count += link(k);
    if (node !== stack[0]) node.dict.set(N('Count'), PDFNumber.of(-node.kids.length)); // collapsed
    return count;
  };
  const total = link(stack[0]);
  root.set(N('Count'), PDFNumber.of(total));
  pdf.catalog.set(N('Outlines'), rootRef);
  pdf.catalog.set(N('PageMode'), N('UseOutlines'));
}
function writeXmp(pdf, title) {
  const esc = (s) => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const now = new Date().toISOString();
  const xmp = `<?xpacket begin="﻿" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about="" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xmp="http://ns.adobe.com/xap/1.0/" xmlns:pdf="http://ns.adobe.com/pdf/1.3/">
   <dc:format>application/pdf</dc:format>
   <dc:title><rdf:Alt><rdf:li xml:lang="x-default">${esc(title)}</rdf:li></rdf:Alt></dc:title>
   ${pdf.getAuthor() ? `<dc:creator><rdf:Seq><rdf:li>${esc(pdf.getAuthor())}</rdf:li></rdf:Seq></dc:creator>` : ''}
   <xmp:ModifyDate>${now}</xmp:ModifyDate>
   <xmp:MetadataDate>${now}</xmp:MetadataDate>
   <pdf:Producer>${esc(pdf.getProducer() || 'Accessible Document Checker')}</pdf:Producer>
  </rdf:Description>
 </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>`;
  const stream = pdf.context.stream(Buffer.from(xmp, 'utf8'), { Type: 'Metadata', Subtype: 'XML' });
  pdf.catalog.set(N('Metadata'), pdf.context.register(stream));
}

/** For AI alt text: a one-page PDF containing the figure's page. */
async function getImage(buffer, target) {
  if (!target || target.page === null || target.page === undefined) return null;
  const src = await loadDoc(buffer, false);
  const out = await PDFDocument.create();
  const [p] = await out.copyPages(src, [target.page]);
  out.addPage(p);
  return { data: Buffer.from(await out.save()), mime: 'application/pdf' };
}

module.exports = { audit, fix, getImage };
