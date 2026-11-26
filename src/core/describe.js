'use strict';
// Offline suggestions that need no AI: alt text for charts from the chart's own
// data, alt text taken from a nearby caption, and link text from the address.
// These are suggestions for a person to review, not final answers.
const X = require('./xml');

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_TYPES = {
  barChart: 'Bar chart', bar3DChart: 'Bar chart', lineChart: 'Line chart', line3DChart: 'Line chart',
  pieChart: 'Pie chart', pie3DChart: 'Pie chart', ofPieChart: 'Pie chart', doughnutChart: 'Doughnut chart',
  areaChart: 'Area chart', area3DChart: 'Area chart', scatterChart: 'Scatter chart', bubbleChart: 'Bubble chart',
  radarChart: 'Radar chart', stockChart: 'Stock chart', surfaceChart: 'Surface chart', surface3DChart: 'Surface chart',
};

const textOf = (node) => (node ? X.all(node, 'a', 't').map((t) => t.textContent).join('').trim() : '');
const kidsC = (node, local) => {
  const out = [];
  for (let c = node && node.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.namespaceURI === C_NS && c.localName === local) out.push(c);
  return out;
};
const firstC = (node, local) => (node ? node.getElementsByTagNameNS(C_NS, local)[0] || null : null);
function points(node) {
  if (!node) return [];
  const pts = Array.from(node.getElementsByTagNameNS(C_NS, 'pt'));
  return pts.sort((a, b) => +a.getAttribute('idx') - +b.getAttribute('idx')).map((p) => (firstC(p, 'v') || {}).textContent || '');
}
function fmt(n) {
  const v = Number(n);
  if (!Number.isFinite(v)) return String(n);
  return Math.abs(v) >= 1000 ? v.toLocaleString('en-GB', { maximumFractionDigits: 1 }) : String(Math.round(v * 100) / 100);
}
const list = (a) => (a.length <= 1 ? a.join('') : `${a.slice(0, -1).join(', ')} and ${a[a.length - 1]}`);

/** Read a DrawingML chart part and return its type, title, categories and series. */
function readChart(xmlText) {
  const doc = X.parse(xmlText);
  const chart = firstC(doc, 'chart');
  if (!chart) return null;
  const titleEl = kidsC(chart, 'title')[0];
  const title = textOf(titleEl);
  const plot = firstC(chart, 'plotArea');
  if (!plot) return null;
  let type = null, horizontal = false;
  const series = [];
  for (let c = plot.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1 || !CHART_TYPES[c.localName]) continue;
    if (!type) {
      type = CHART_TYPES[c.localName];
      const dir = firstC(c, 'barDir');
      if (/^bar/.test(c.localName)) { horizontal = dir && dir.getAttribute('val') === 'bar'; type = horizontal ? 'Bar chart' : 'Column chart'; }
    }
    for (const ser of kidsC(c, 'ser')) {
      const name = (firstC(kidsC(ser, 'tx')[0], 'v') || {}).textContent || textOf(kidsC(ser, 'tx')[0]) || '';
      const cats = points(kidsC(ser, 'cat')[0] || kidsC(ser, 'xVal')[0]);
      const vals = points(kidsC(ser, 'val')[0] || kidsC(ser, 'yVal')[0]);
      series.push({ name: name.trim(), cats, vals });
    }
  }
  if (!type) return null;
  const axisTitles = Array.from(plot.getElementsByTagNameNS(C_NS, 'title')).map(textOf).filter(Boolean);
  return { type, title, series, axisTitles };
}

/** Plain-language alt text from chart data, e.g. "Column chart: Sales by region. North 40, South 25. Highest: North (40)." */
function describeChart(xmlText) {
  let c;
  try { c = readChart(xmlText); } catch { return ''; }
  if (!c) return '';
  const bits = [c.title ? `${c.type}: ${c.title}.` : `${c.type}.`];
  const pie = /Pie|Doughnut/.test(c.type);
  if (c.axisTitles.length && !pie) bits.push(`Axes: ${list(c.axisTitles)}.`);
  const many = c.series.length > 1;
  for (const s of c.series.slice(0, 4)) {
    const pairs = s.vals.map((v, i) => ({ cat: s.cats[i] || `Item ${i + 1}`, v: Number(v), raw: v })).filter((p) => p.raw !== '');
    if (!pairs.length) continue;
    const label = many && s.name ? `${s.name}: ` : '';
    if (pairs.length <= 6) {
      bits.push(`${label}${pairs.map((p) => `${p.cat} ${fmt(p.raw)}`).join(', ')}.`);
    } else {
      bits.push(`${label}${pairs.length} values from ${pairs[0].cat} to ${pairs[pairs.length - 1].cat}.`);
    }
    const nums = pairs.filter((p) => Number.isFinite(p.v));
    if (nums.length >= 3) {
      const hi = nums.reduce((a, b) => (b.v > a.v ? b : a));
      const lo = nums.reduce((a, b) => (b.v < a.v ? b : a));
      if (c.type === 'Line chart' || c.type === 'Area chart') {
        const dir = nums[nums.length - 1].v > nums[0].v ? 'rises' : nums[nums.length - 1].v < nums[0].v ? 'falls' : 'ends level';
        bits.push(`${many && s.name ? s.name : 'It'} ${dir} from ${fmt(nums[0].v)} to ${fmt(nums[nums.length - 1].v)}; highest ${hi.cat} (${fmt(hi.v)}).`);
      } else {
        bits.push(`Highest: ${hi.cat} (${fmt(hi.v)}). Lowest: ${lo.cat} (${fmt(lo.v)}).`);
      }
    }
  }
  if (c.series.length > 4) bits.push(`Plus ${c.series.length - 4} more series.`);
  let out = bits.join(' ');
  if (out.length > 400) out = out.slice(0, 397).replace(/\s+\S*$/, '') + '…';
  return out;
}

// "Figure 2: Map of offices", "Chart 1 – Sales", "Photo: Students in class"
const CAPTION = /^\s*(figure|fig\.?|chart|graph|image|photo(graph)?|picture|diagram|map|illustration|infographic)\s*[\dIVX.]*\s*[:.\-–—]\s*(.{3,})$/i;

/** Text of a caption line, without the "Figure 1:" label, or ''. */
function captionText(line) {
  const m = String(line || '').trim().match(CAPTION);
  if (!m) return '';
  return m[3].trim().replace(/\s+/g, ' ').slice(0, 250);
}
/** First caption found in a list of candidate lines (nearest first). */
function findCaption(lines) {
  for (const l of lines) { const c = captionText(l); if (c) return c; }
  return '';
}

module.exports = { describeChart, readChart, captionText, findCaption, CAPTION };
