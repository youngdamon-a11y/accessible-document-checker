'use strict';
// Small helpers around @xmldom/xmldom for editing Office Open XML parts.
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  adec: 'http://schemas.microsoft.com/office/drawing/2017/decorative',
};
const DECORATIVE_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';

function parse(text) {
  return new DOMParser({ onError: () => {} }).parseFromString(text, 'text/xml');
}
function serialize(doc) {
  let s = new XMLSerializer().serializeToString(doc);
  if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  return s;
}
/** All descendants with namespace + local name, as an array. */
function all(node, ns, local) {
  return Array.from(node.getElementsByTagNameNS(NS[ns] || ns, local));
}
function first(node, ns, local) {
  return node.getElementsByTagNameNS(NS[ns] || ns, local)[0] || null;
}
/** Direct children matching ns:local. */
function kids(node, ns, local) {
  const out = [];
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && c.namespaceURI === (NS[ns] || ns) && c.localName === local) out.push(c);
  }
  return out;
}
function kid(node, ns, local) {
  return kids(node, ns, local)[0] || null;
}
function attr(el, ns, local) {
  if (!el) return null;
  if (!ns) return el.hasAttribute(local) ? el.getAttribute(local) : null;
  const v = el.getAttributeNS(NS[ns] || ns, local);
  return v === '' && !el.hasAttributeNS(NS[ns] || ns, local) ? null : v;
}
function setAttr(el, ns, qname, value) {
  if (!ns) el.setAttribute(qname, value);
  else el.setAttributeNS(NS[ns] || ns, qname, value);
}
/** Create an element in a namespace using the conventional prefix. */
function el(doc, ns, local, attrs) {
  const e = doc.createElementNS(NS[ns] || ns, `${ns}:${local}`);
  if (attrs) for (const [k, v] of Object.entries(attrs)) {
    if (k.includes(':')) { const [p] = k.split(':'); e.setAttributeNS(NS[p], k, v); } else e.setAttribute(k, v);
  }
  return e;
}
/** Get or create a direct child, inserting as the first child when insertFirst. */
function ensureKid(parent, ns, local, insertFirst) {
  let c = kid(parent, ns, local);
  if (!c) {
    c = el(parent.ownerDocument, ns, local);
    if (insertFirst && parent.firstChild) parent.insertBefore(c, parent.firstChild);
    else parent.appendChild(c);
  }
  return c;
}
function text(node) {
  return node ? (node.textContent || '') : '';
}
function remove(node) {
  if (node && node.parentNode) node.parentNode.removeChild(node);
}

/** Is a drawing marked decorative (Office 2019+ extension)? */
function isDecorative(cNvPr) {
  if (!cNvPr) return false;
  for (const ext of all(cNvPr, 'a', 'ext')) {
    if (ext.getAttribute('uri') === DECORATIVE_URI) {
      const d = first(ext, 'adec', 'decorative');
      if (d && d.getAttribute('val') === '1') return true;
    }
  }
  return false;
}
function markDecorative(cNvPr) {
  const doc = cNvPr.ownerDocument;
  cNvPr.setAttribute('descr', '');
  if (isDecorative(cNvPr)) return;
  let extLst = kid(cNvPr, 'a', 'extLst');
  if (!extLst) { extLst = el(doc, 'a', 'extLst'); cNvPr.appendChild(extLst); }
  const ext = el(doc, 'a', 'ext', { uri: DECORATIVE_URI });
  const dec = doc.createElementNS(NS.adec, 'adec:decorative');
  dec.setAttribute('val', '1');
  ext.appendChild(dec);
  extLst.appendChild(ext);
}
function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

module.exports = { NS, DECORATIVE_URI, parse, serialize, all, first, kids, kid, attr, setAttr, el, ensureKid, text, remove, isDecorative, markDecorative, escapeXml };
