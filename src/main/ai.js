'use strict';
// Claude suggestions for alt text, link text and slide titles.
// Only used when the person asks for a suggestion; everything else works offline.
const Anthropic = require('@anthropic-ai/sdk');

const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
const MAX_IMAGE_BYTES = 3.5 * 1024 * 1024;

const SYSTEM = `You write accessibility text for documents, following WCAG 2.1 and UK government / British English style.
Rules:
- Reply with only the requested text. No quotes, no preamble, no explanation.
- Alt text: one or two concise sentences (usually under 150 characters). Convey the purpose and key information in context, not just appearance. Do not start with "Image of" or "Picture of". For charts, give the type and the main finding or key figures. If the image contains important text, include that text.
- If an image is purely decorative and adds no information, reply exactly: DECORATIVE
- Link text: a short phrase (2 to 8 words) that says where the link goes, so it makes sense out of context.
- Slide titles: short and specific (under 10 words).`;

function client(apiKey) {
  if (!apiKey) throw new Error('Add your Claude API key in Settings to get suggestions. You can still type text yourself.');
  return new Anthropic({ apiKey, timeout: 90_000, maxRetries: 2 });
}

/** Ensure the image is a type and size Claude accepts; convert with nativeImage if needed. */
function prepareImage(img) {
  let { data, mime } = img;
  if (mime === 'application/pdf') return { kind: 'document', data, mime };
  if (!IMAGE_TYPES.includes(mime) || data.length > MAX_IMAGE_BYTES) {
    const { nativeImage } = require('electron');
    let ni = nativeImage.createFromBuffer(data);
    if (ni.isEmpty()) throw new Error(`This image format (${mime.replace('image/', '')}) cannot be sent for a suggestion. Please type the alt text yourself.`);
    const { width } = ni.getSize();
    if (width > 1600) ni = ni.resize({ width: 1600, quality: 'good' });
    data = ni.toJPEG(85); mime = 'image/jpeg';
  }
  return { kind: 'image', data, mime };
}

async function ask(apiKey, model, content) {
  const c = client(apiKey);
  const params = {
    model: model || 'claude-opus-5-5',
    max_tokens: 4000,
    system: SYSTEM,
    output_config: { effort: 'low' },
    messages: [{ role: 'user', content }],
  };
  // Server-side fallback re-runs a declined request on another model.
  if (!/haiku/.test(params.model)) { params.betas = ['server-side-fallback-2026-07-01']; params.fallbacks = 'default'; }
  let res;
  try {
    res = await c.beta.messages.create(params);
  } catch (e) {
    if (e instanceof Anthropic.AuthenticationError) throw new Error('Claude rejected the API key. Check it in Settings.');
    if (e instanceof Anthropic.RateLimitError) throw new Error('Claude is rate limiting requests. Wait a moment and try again.');
    if (e instanceof Anthropic.APIConnectionError) throw new Error('Could not reach Claude. Check your internet connection (everything else works offline).');
    if (e instanceof Anthropic.APIError) throw new Error(`Claude returned an error: ${e.message}`);
    throw e;
  }
  if (res.stop_reason === 'refusal') throw new Error('Claude declined to describe this item. Please write the text yourself.');
  const text = res.content.filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
  if (!text) throw new Error('Claude did not return a suggestion. Try again or type it yourself.');
  return text.replace(/^["“']|["”']$/g, '').trim();
}

async function suggestAltText({ apiKey, model, image, context, language, figureHint }) {
  const prepared = prepareImage(image);
  const content = [];
  if (prepared.kind === 'document') {
    content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: prepared.data.toString('base64') } });
  } else {
    content.push({ type: 'image', source: { type: 'base64', media_type: prepared.mime, data: prepared.data.toString('base64') } });
  }
  content.push({
    type: 'text',
    text: `${prepared.kind === 'document' ? `This is one page of a PDF. Write alt text for ${figureHint || 'the main figure or image'} on this page.` : 'Write alt text for this image.'}
Write it in the document's language (${language || 'en-GB'}).
Surrounding document text, for context:
"""
${(context || '').slice(0, 1500)}
"""`,
  });
  const text = await ask(apiKey, model, content);
  return text.toUpperCase() === 'DECORATIVE' ? { decorative: true, text: '' } : { decorative: false, text };
}

async function suggestText({ apiKey, model, kind, current, url, context, language }) {
  const what = kind === 'slide-title' ? 'a slide title for this slide' : `descriptive link text to replace "${current}" (link goes to ${url || 'unknown address'})`;
  const text = await ask(apiKey, model, [{ type: 'text', text: `Write ${what}. Language: ${language || 'en-GB'}.\nContext:\n"""\n${(context || '').slice(0, 1500)}\n"""` }]);
  return { text };
}

async function testKey(apiKey, model) {
  const c = client(apiKey);
  try {
    const m = await c.models.retrieve(model || 'claude-opus-5-5');
    return { ok: true, model: m.display_name || m.id };
  } catch (e) {
    if (e instanceof Anthropic.AuthenticationError) return { ok: false, message: 'The API key was rejected.' };
    if (e instanceof Anthropic.NotFoundError) return { ok: false, message: 'The key works, but this model is not available to your account.' };
    if (e instanceof Anthropic.APIConnectionError) return { ok: false, message: 'Could not reach Claude. Check your internet connection.' };
    return { ok: false, message: e.message };
  }
}

module.exports = { suggestAltText, suggestText, testKey };
