'use strict';
// Builds the exportable accessibility report (accessible HTML; the app can also print it to a tagged PDF).

const esc = (s) => String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * @param {object} d
 *  fileName, savedPath, format, original (audit result before fixes), current (latest audit),
 *  applied [{title, wcag, at, by}], humanChecks {id: {status:'confirmed'|'not-applicable'|'', note}}
 */
function buildReport(d) {
  const cur = d.current;
  const before = d.original;
  const humanList = cur.humanChecks || [];
  const hcState = d.humanChecks || {};
  const outstandingHuman = humanList.filter((h) => !['confirmed', 'not-applicable'].includes((hcState[h.id] || {}).status));
  const errors = cur.issues.filter((i) => i.severity === 'error');
  const warnings = cur.issues.filter((i) => i.severity !== 'error');
  let verdict, verdictClass;
  if (errors.length) { verdict = `Does not pass: ${errors.length} problem${errors.length > 1 ? 's' : ''} still to fix`; verdictClass = 'fail'; }
  else if (outstandingHuman.length) { verdict = `Passes automated checks. ${outstandingHuman.length} human check${outstandingHuman.length > 1 ? 's' : ''} still to confirm`; verdictClass = 'partial'; }
  else { verdict = 'Passes automated checks and all human checks are confirmed'; verdictClass = 'pass'; }

  const crit = (i) => i.wcag.map((c) => `${esc(c.id)} ${esc(c.name)} (${esc(c.level)})`).join('; ') + (i.pdfua ? `; PDF/UA ${esc(i.pdfua)}` : '');
  const issueRows = (list) => list.map((i) => `<tr><td>${esc(i.title)}${i.description ? `<br><span class="muted">${esc(i.description)}</span>` : ''}</td><td>${crit(i)}</td><td>${esc(i.location.label)}</td><td>${{ auto: 'Automatic', ai: 'Needs text (AI can suggest)', input: 'Needs your input', manual: 'Fix in source / manually' }[i.fix.kind]}</td></tr>`).join('');
  const table = (caption, list) => list.length ? `<table><caption>${esc(caption)}</caption><thead><tr><th scope="col">Issue</th><th scope="col">Criteria</th><th scope="col">Location</th><th scope="col">Fix type</th></tr></thead><tbody>${issueRows(list)}</tbody></table>` : `<p>None.</p>`;

  const appliedRows = (d.applied || []).map((a) => `<li>${a.suggested ? '<strong>Review:</strong> ' : ''}${esc(a.title)}${a.detail ? ` — ${esc(a.detail)}` : ''}${a.by ? ` <span class="muted">(${esc(a.by)})</span>` : ''}</li>`).join('');
  const suggestedCount = (d.applied || []).filter((a) => a.suggested).length;
  const hcRows = humanList.map((h) => {
    const st = hcState[h.id] || {};
    const label = st.status === 'confirmed' ? 'Confirmed' : st.status === 'not-applicable' ? 'Not applicable' : 'Not yet checked';
    const details = (cur.humanCheckDetails || {})[h.id];
    return `<tr><td>${esc(h.title)}<br><span class="muted">${esc(h.description)}</span>${details ? `<details><summary>Details</summary><ul>${details.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></details>` : ''}</td><td>${h.wcag.map((c) => `${esc(c.id)} ${esc(c.name)}`).join('; ')}</td><td>${label}${st.note ? `<br><span class="muted">${esc(st.note)}</span>` : ''}</td></tr>`;
  }).join('');

  const metaRows = Object.entries(cur.meta || {}).filter(([, v]) => v !== null && v !== undefined).map(([k, v]) => `<tr><th scope="row">${esc(k.charAt(0).toUpperCase() + k.slice(1))}</th><td>${esc(typeof v === 'boolean' ? (v ? 'Yes' : 'No') : v)}</td></tr>`).join('');
  const date = new Date(cur.checkedAt || Date.now());
  const title = `Accessibility report: ${d.fileName}`;
  return `<!doctype html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
  body { font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif; color: #1A1A2E; margin: 2rem auto; max-width: 60rem; padding: 0 1rem; line-height: 1.5; }
  h1 { color: #23085A; font-size: 1.8rem; margin-bottom: .25rem; }
  h2 { color: #23085A; border-bottom: 3px solid #23085A; padding-bottom: .25rem; margin-top: 2rem; }
  .muted { color: #4A4A5E; font-size: .92em; }
  .verdict { padding: 1rem 1.25rem; border-radius: 6px; font-weight: 700; font-size: 1.15rem; border-left: 8px solid; }
  .pass { background: #E6F4EA; border-color: #0B6E35; color: #0B4D26; }
  .partial { background: #FFF4D6; border-color: #8A5A00; color: #5C3D00; }
  .fail { background: #FDE8EA; border-color: #B00020; color: #7A0016; }
  table { border-collapse: collapse; width: 100%; margin: 1rem 0; font-size: .95rem; }
  caption { text-align: left; font-weight: 700; padding: .5rem 0; }
  th, td { border: 1px solid #8C8C9E; padding: .5rem; text-align: left; vertical-align: top; }
  thead th { background: #23085A; color: #fff; }
  th[scope=row] { background: #F3F2F7; width: 30%; }
  .counts { display: flex; gap: 1rem; flex-wrap: wrap; padding: 0; list-style: none; }
  .counts li { background: #F3F2F7; padding: .75rem 1rem; border-radius: 6px; }
  .counts strong { display: block; font-size: 1.5rem; color: #23085A; }
  footer { margin-top: 3rem; font-size: .9rem; color: #4A4A5E; }
  @media print { body { margin: 0; } h2 { break-after: avoid; } tr { break-inside: avoid; } }
</style>
</head>
<body>
<header>
  <h1>${esc(title)}</h1>
  <p class="muted">Checked against WCAG 2.1 level AA${d.format === 'pdf' ? ' and key PDF/UA (ISO 14289-1) requirements' : ''} on ${esc(date.toLocaleString('en-GB'))}.${d.savedPath ? ` Remediated copy: ${esc(d.savedPath)}` : ''}</p>
</header>
<main>
  <p class="verdict ${verdictClass}" role="status">${esc(verdict)}</p>
  <h2>Summary</h2>
  <ul class="counts">
    <li><strong>${before ? before.issues.length : '–'}</strong> issues found at first check</li>
    <li><strong>${(d.applied || []).length}</strong> fixes applied</li>
    <li><strong>${errors.length}</strong> errors remaining</li>
    <li><strong>${warnings.length}</strong> warnings remaining</li>
    <li><strong>${outstandingHuman.length}</strong> human checks outstanding</li>
  </ul>
  <table><caption>Document properties</caption><tbody><tr><th scope="row">File</th><td>${esc(d.fileName)}</td></tr>${metaRows}</tbody></table>
  <h2>Remaining errors</h2>
  ${table('Errors that still fail the checks', errors)}
  <h2>Remaining warnings</h2>
  ${table('Warnings to review', warnings)}
  <h2>Fixes applied</h2>
  ${suggestedCount ? `<p>${suggestedCount} item${suggestedCount > 1 ? 's were' : ' was'} filled in from suggested text (captions, chart data, web addresses or an AI model) and should be checked by a person. They are marked “Review”.</p>` : ''}
  ${appliedRows ? `<ol>${appliedRows}</ol>` : '<p>No fixes applied.</p>'}
  <h2>Human checks</h2>
  <p>Automated tools cannot judge every WCAG criterion. These need a person to confirm.</p>
  <table><caption>Checks that need human judgement</caption><thead><tr><th scope="col">Check</th><th scope="col">Criteria</th><th scope="col">Status</th></tr></thead><tbody>${hcRows}</tbody></table>
  ${before ? `<h2>Issues found at first check</h2>${table('All issues found before fixing', before.issues)}` : ''}
</main>
<footer>
  <p><strong>About this report.</strong> Automated checks cover the parts of WCAG 2.1 AA that can be tested by software. Passing them does not by itself guarantee full conformance: reading order, the quality of alt text and link text, use of colour and other items listed under “Human checks” need a person to confirm. ${d.format === 'pdf' ? 'For formal PDF/UA conformance, also validate with a dedicated tool such as PAC or veraPDF.' : ''}</p>
  <p>Generated by Accessible Document Checker.</p>
</footer>
</body>
</html>`;
}

module.exports = { buildReport };
