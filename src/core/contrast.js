'use strict';
// WCAG 2.1 contrast maths (SC 1.4.3) and a colour adjuster that keeps the hue.

function hexToRgb(hex) {
  const h = String(hex).replace('#', '').trim();
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return null;
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
}
function rgbToHex([r, g, b]) {
  return [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}
function channel(v) {
  const s = v / 255;
  return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
}
function luminance(rgb) {
  return 0.2126 * channel(rgb[0]) + 0.7152 * channel(rgb[1]) + 0.0722 * channel(rgb[2]);
}
function ratio(fgHex, bgHex) {
  const a = hexToRgb(fgHex), b = hexToRgb(bgHex);
  if (!a || !b) return null;
  const l1 = luminance(a), l2 = luminance(b);
  return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
}
/** Large text per WCAG: >= 18pt, or >= 14pt bold. */
function requiredRatio(sizePt, bold) {
  if (sizePt && (sizePt >= 18 || (bold && sizePt >= 14))) return 3;
  return 4.5;
}
/**
 * Return a new foreground hex that meets `target` against bg, moving the colour
 * towards black or white (whichever direction gives more contrast) in small steps
 * so the original hue is kept as far as possible.
 */
function fixForeground(fgHex, bgHex, target = 4.5) {
  const fg = hexToRgb(fgHex), bg = hexToRgb(bgHex);
  if (!fg || !bg) return null;
  const towards = luminance(bg) > 0.18 ? [0, 0, 0] : [255, 255, 255];
  for (let t = 0; t <= 1.0001; t += 0.02) {
    const c = fg.map((v, i) => v + (towards[i] - v) * t);
    const hex = rgbToHex(c);
    if (ratio(hex, bgHex) >= target + 0.05) return hex;
  }
  return rgbToHex(towards);
}

module.exports = { hexToRgb, rgbToHex, luminance, ratio, requiredRatio, fixForeground };
