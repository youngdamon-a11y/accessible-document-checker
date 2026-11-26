'use strict';
// Shared issue model, WCAG criteria catalogue and the human-check list.

const WCAG = {
  '1.1.1': { name: 'Non-text Content', level: 'A' },
  '1.2.2': { name: 'Captions (Prerecorded)', level: 'A' },
  '1.3.1': { name: 'Info and Relationships', level: 'A' },
  '1.3.2': { name: 'Meaningful Sequence', level: 'A' },
  '1.3.3': { name: 'Sensory Characteristics', level: 'A' },
  '1.4.1': { name: 'Use of Color', level: 'A' },
  '1.4.3': { name: 'Contrast (Minimum)', level: 'AA' },
  '1.4.5': { name: 'Images of Text', level: 'AA' },
  '2.4.2': { name: 'Page Titled', level: 'A' },
  '2.4.3': { name: 'Focus Order', level: 'A' },
  '2.4.4': { name: 'Link Purpose (In Context)', level: 'A' },
  '2.4.5': { name: 'Multiple Ways', level: 'AA' },
  '2.4.6': { name: 'Headings and Labels', level: 'AA' },
  '3.1.1': { name: 'Language of Page', level: 'A' },
  '4.1.2': { name: 'Name, Role, Value', level: 'A' },
};

/**
 * Fix kinds:
 *  auto   - the app can fix it with no input (included in "Fix all")
 *  ai     - needs text; Claude can suggest it, the user approves or edits it
 *  input  - needs text from the user (an offline suggestion may be pre-filled)
 *  manual - cannot be fixed automatically; the user must check or fix it in the source
 */
function makeIssue(o) {
  const crit = (o.wcag || []).map((id) => ({ id, ...(WCAG[id] || { name: '', level: '' }) }));
  return {
    id: o.id,
    rule: o.rule,
    severity: o.severity || 'error', // error | warning
    title: o.title,
    description: o.description,
    wcag: crit,
    pdfua: o.pdfua || null,
    location: o.location || { label: 'Whole document' },
    fix: o.fix || { kind: 'manual' },
    target: o.target || null,
  };
}

const HUMAN_CHECKS = {
  common: [
    { id: 'hc-reading-order', wcag: ['1.3.2'], title: 'Reading order makes sense', description: 'Read through the document in order (or with a screen reader) and confirm content is announced in a logical sequence, including text boxes, sidebars and columns.' },
    { id: 'hc-alt-quality', wcag: ['1.1.1'], title: 'Alt text is accurate and useful', description: 'Review every alt text, especially AI suggestions. It should convey the purpose of the image in context, not just describe it. Charts need their key message or data.' },
    { id: 'hc-colour', wcag: ['1.4.1'], title: 'Colour is not the only way information is shown', description: 'Check charts, highlights, status labels and links do not rely on colour alone (add text, patterns or underlines).' },
    { id: 'hc-sensory', wcag: ['1.3.3'], title: 'Instructions do not rely on shape, position or sound', description: 'Phrases like "click the green button" or "see the box on the right" need another cue, such as a name or heading.' },
    { id: 'hc-images-of-text', wcag: ['1.4.5'], title: 'No images of text', description: 'Text should be real text, not a picture of text (except logos). Screenshots of tables or paragraphs should be replaced.' },
    { id: 'hc-link-context', wcag: ['2.4.4'], title: 'Link text makes sense', description: 'Each link should say where it goes. Check links the app could not judge, and any link text you changed.' },
  ],
  pptx: [
    { id: 'hc-media', wcag: ['1.2.2'], title: 'Video and audio have captions or transcripts', description: 'If the presentation contains media, confirm captions or a transcript are provided.' },
  ],
  pdf: [
    { id: 'hc-pdf-tags', wcag: ['1.3.1', '1.3.2'], title: 'Tags match the visible content', description: 'Open the tag tree (e.g. in Acrobat or PAC) and confirm headings, lists, tables and figures are tagged correctly and in order.' },
  ],
};

function humanChecks(format) {
  return [...HUMAN_CHECKS.common, ...(HUMAN_CHECKS[format] || [])].map((h) => ({
    ...h,
    wcag: h.wcag.map((id) => ({ id, ...WCAG[id] })),
  }));
}

const LINK_TEXT_BAD = /^(click here|here|click|more|read more|learn more|link|this link|this|go|details|more info|more information|download|website|page)\.?$/i;
function isUrlText(t) {
  return /^(https?:\/\/|www\.|mailto:)/i.test(t.trim()) || /^[\w.-]+\.(com|org|net|gov|uk|edu|io|co)(\/\S*)?$/i.test(t.trim());
}
/** Offline fallback suggestion for link text from a URL. */
function suggestLinkText(url, current) {
  try {
    if (/^mailto:/i.test(url)) return `Email ${url.replace(/^mailto:/i, '').split('?')[0]}`;
    const u = new URL(/^https?:/i.test(url) ? url : 'https://' + url);
    const host = u.hostname.replace(/^www\./, '');
    const last = u.pathname.split('/').filter(Boolean).pop();
    if (last) {
      const words = decodeURIComponent(last).replace(/\.[a-z0-9]{2,5}$/i, '').replace(/[-_]+/g, ' ').trim();
      if (words && !/^\d+$/.test(words)) return `${words.charAt(0).toUpperCase()}${words.slice(1)} (${host})`;
    }
    return `${host} website`;
  } catch {
    return current || 'Descriptive link text';
  }
}

module.exports = { WCAG, makeIssue, humanChecks, LINK_TEXT_BAD, isUrlText, suggestLinkText };
