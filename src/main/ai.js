'use strict';
// AI suggestions for alt text, link text and slide titles. Only used when the person
// asks for a suggestion; everything else works offline. Three providers:
//   anthropic  Claude API, with the person's own API key
//   openai     OpenAI API, with the person's own API key
//   ollama     a vision model running on this computer through Ollama (fully offline)
// Subscription sign-in (Claude Pro/Max, ChatGPT Plus/Pro) is not offered: see README.
const Anthropic = require('@anthropic-ai/sdk');

const IMAGE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'];
const MAX_IMAGE_BYTES = 3.5 * 1024 * 1024;

const PROVIDERS = {
  anthropic: { name: 'Claude', defaultModel: 'claude-opus-5-5', needsKey: true },
  openai: { name: 'OpenAI', defaultModel: 'gpt-5', needsKey: true },
  ollama: { name: 'Local model (Ollama)', defaultModel: 'gemma3', needsKey: false },
};

const SYSTEM = `You write accessibility text for documents, following WCAG 2.1 and UK government / British English style.
Rules:
- Reply with only the requested text. No quotes, no preamble, no explanation.
- Alt text: one or two concise sentences (usually under 150 characters). Convey the purpose and key information in context, not just appearance. Do not start with "Image of" or "Picture of". For charts, give the type and the main finding or key figures. If the image contains important text, include that text.
- If an image is purely decorative and adds no information, reply exactly: DECORATIVE
- Link text: a short phrase (2 to 8 words) that says where the link goes, so it makes sense out of context.
- Slide titles: short and specific (under 10 words).`;

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------
/**
 * Make the image something the provider accepts.
 * Claude and OpenAI read a one-page PDF directly; Ollama needs a picture, so PDF pages
 * are rendered with `renderPdf` (supplied by the main process).
 */
async function prepareImage(img, provider, renderPdf) {
  let { data, mime } = img;
  if (mime === 'application/pdf') {
    if (provider !== 'ollama') return { kind: 'document', data, mime };
    if (!renderPdf) throw new Error('This PDF page could not be turned into a picture for the local model.');
    data = await renderPdf(data); mime = 'image/png';
  }
  const okTypes = provider === 'ollama' ? ['image/png', 'image/jpeg'] : IMAGE_TYPES;
  if (!okTypes.includes(mime) || data.length > MAX_IMAGE_BYTES) {
    const { nativeImage } = require('electron');
    let ni = nativeImage.createFromBuffer(data);
    if (ni.isEmpty()) throw new Error(`This image format (${mime.replace('image/', '')}) cannot be sent for a suggestion. Please type the alt text yourself.`);
    const { width } = ni.getSize();
    if (width > 1600) ni = ni.resize({ width: 1600, quality: 'good' });
    data = ni.toJPEG(85); mime = 'image/jpeg';
  }
  return { kind: 'image', data, mime };
}

const clean = (t) => String(t || '').replace(/<think>[\s\S]*?<\/think>/gi, '').trim().replace(/^["“']|["”']$/g, '').trim();

// ---------------------------------------------------------------------------
// Providers. Each takes { prompt, image } and returns plain text.
// ---------------------------------------------------------------------------
async function askAnthropic(cfg, { prompt, image }) {
  if (!cfg.apiKey) throw new Error('Add your Claude API key in Settings to get suggestions. You can still type text yourself.');
  const c = new Anthropic({ apiKey: cfg.apiKey, timeout: 90_000, maxRetries: 2 });
  const content = [];
  if (image && image.kind === 'document') content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: image.data.toString('base64') } });
  else if (image) content.push({ type: 'image', source: { type: 'base64', media_type: image.mime, data: image.data.toString('base64') } });
  content.push({ type: 'text', text: prompt });
  const params = {
    model: cfg.model || PROVIDERS.anthropic.defaultModel,
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
  return res.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
}

async function fetchJson(url, init, who, timeoutMs = 120_000) {
  let res;
  try {
    res = await fetch(url, { ...init, signal: AbortSignal.timeout(timeoutMs) });
  } catch (e) {
    if (who === 'ollama') throw new Error('Could not reach Ollama on this computer. Check it is installed and running, then select Test connection in Settings.');
    throw new Error(`Could not reach ${who === 'openai' ? 'OpenAI' : who}. Check your internet connection (everything else works offline).`);
  }
  let body = null;
  try { body = await res.json(); } catch { /* not JSON */ }
  if (!res.ok) {
    const msg = (body && (body.error && (body.error.message || body.error))) || res.statusText;
    if (who === 'openai' && res.status === 401) throw new Error('OpenAI rejected the API key. Check it in Settings.');
    if (who === 'openai' && res.status === 429) throw new Error('OpenAI is rate limiting requests, or your account has no credit left. Wait a moment and try again.');
    if (who === 'ollama' && res.status === 404) throw new Error(`The local model is not installed. In a terminal run: ollama pull ${init.model || 'gemma3'}`);
    throw new Error(`${who === 'openai' ? 'OpenAI' : 'Ollama'} returned an error: ${msg}`);
  }
  return body;
}

async function askOpenAI(cfg, { prompt, image }) {
  if (!cfg.apiKey) throw new Error('Add your OpenAI API key in Settings to get suggestions. You can still type text yourself.');
  const content = [];
  if (image && image.kind === 'document') content.push({ type: 'input_file', filename: 'page.pdf', file_data: `data:application/pdf;base64,${image.data.toString('base64')}` });
  else if (image) content.push({ type: 'input_image', image_url: `data:${image.mime};base64,${image.data.toString('base64')}` });
  content.push({ type: 'input_text', text: prompt });
  const body = await fetchJson('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: { 'content-type': 'application/json', authorization: `Bearer ${cfg.apiKey}` },
    body: JSON.stringify({ model: cfg.model || PROVIDERS.openai.defaultModel, instructions: SYSTEM, input: [{ role: 'user', content }], store: false }),
  }, 'openai');
  const out = [];
  for (const item of body.output || []) {
    if (item.type === 'message') for (const c of item.content || []) {
      if (c.type === 'output_text') out.push(c.text);
      if (c.type === 'refusal') throw new Error('OpenAI declined to describe this item. Please write the text yourself.');
    }
  }
  return out.join('');
}

function ollamaBase(cfg) {
  const u = new URL(cfg.url || 'http://localhost:11434');
  if (!/^https?:$/.test(u.protocol)) throw new Error('The Ollama address must start with http:// or https://');
  return u.origin;
}
async function askOllama(cfg, { prompt, image }) {
  const model = cfg.model || PROVIDERS.ollama.defaultModel;
  const msg = { role: 'user', content: prompt };
  if (image) msg.images = [image.data.toString('base64')];
  const body = await fetchJson(`${ollamaBase(cfg)}/api/chat`, {
    method: 'POST', model,
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ model, stream: false, think: false, options: { temperature: 0.2 }, messages: [{ role: 'system', content: SYSTEM }, msg] }),
  }, 'ollama', 300_000);
  return (body && body.message && body.message.content) || '';
}

const ASK = { anthropic: askAnthropic, openai: askOpenAI, ollama: askOllama };

async function ask(cfg, req) {
  const fn = ASK[cfg.provider];
  if (!fn) throw new Error('AI suggestions are turned off in Settings.');
  const text = clean(await fn(cfg, req));
  if (!text) throw new Error('No suggestion came back. Try again or type it yourself.');
  return text;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------
async function suggestAltText({ cfg, image, context, language, figureHint, renderPdf }) {
  const prepared = await prepareImage(image, cfg.provider, renderPdf);
  const pdfPage = image.mime === 'application/pdf';
  const prompt = `${pdfPage ? `This is one page of a PDF. Write alt text for ${figureHint || 'the main figure or image'} on this page.` : 'Write alt text for this image.'}
Write it in the document's language (${language || 'en-GB'}).
Surrounding document text, for context:
"""
${(context || '').slice(0, 1500)}
"""`;
  const text = await ask(cfg, { prompt, image: prepared });
  return text.toUpperCase().replace(/[.\s]/g, '') === 'DECORATIVE' ? { decorative: true, text: '' } : { decorative: false, text };
}

async function suggestText({ cfg, kind, current, url, context, language }) {
  const what = kind === 'slide-title' ? 'a slide title for this slide' : `descriptive link text to replace "${current}" (link goes to ${url || 'unknown address'})`;
  const text = await ask(cfg, { prompt: `Write ${what}. Language: ${language || 'en-GB'}.\nContext:\n"""\n${(context || '').slice(0, 1500)}\n"""` });
  return { text };
}

/** Check the provider is reachable and the key/model works. Returns { ok, model?, message? }. */
async function testConnection(cfg) {
  try {
    if (cfg.provider === 'anthropic') {
      if (!cfg.apiKey) return { ok: false, message: 'Enter a Claude API key first.' };
      const c = new Anthropic({ apiKey: cfg.apiKey, timeout: 30_000, maxRetries: 1 });
      try {
        const m = await c.models.retrieve(cfg.model || PROVIDERS.anthropic.defaultModel);
        return { ok: true, model: m.display_name || m.id };
      } catch (e) {
        if (e instanceof Anthropic.AuthenticationError) return { ok: false, message: 'The API key was rejected.' };
        if (e instanceof Anthropic.NotFoundError) return { ok: false, message: 'The key works, but this model is not available to your account.' };
        if (e instanceof Anthropic.APIConnectionError) return { ok: false, message: 'Could not reach Claude. Check your internet connection.' };
        return { ok: false, message: e.message };
      }
    }
    if (cfg.provider === 'openai') {
      if (!cfg.apiKey) return { ok: false, message: 'Enter an OpenAI API key first.' };
      const model = cfg.model || PROVIDERS.openai.defaultModel;
      const res = await fetch(`https://api.openai.com/v1/models/${encodeURIComponent(model)}`, { headers: { authorization: `Bearer ${cfg.apiKey}` }, signal: AbortSignal.timeout(30_000) }).catch(() => null);
      if (!res) return { ok: false, message: 'Could not reach OpenAI. Check your internet connection.' };
      if (res.status === 401) return { ok: false, message: 'The API key was rejected.' };
      if (res.status === 404) return { ok: false, message: `The key works, but the model “${model}” is not available to your account.` };
      if (!res.ok) return { ok: false, message: `OpenAI returned ${res.status}.` };
      return { ok: true, model };
    }
    if (cfg.provider === 'ollama') {
      const model = cfg.model || PROVIDERS.ollama.defaultModel;
      const res = await fetch(`${ollamaBase(cfg)}/api/tags`, { signal: AbortSignal.timeout(10_000) }).catch(() => null);
      if (!res || !res.ok) return { ok: false, message: 'Ollama is not running on this computer. Install it from ollama.com, open it, then try again.' };
      const { models = [] } = await res.json();
      const names = models.map((m) => m.name);
      const found = names.some((n) => n === model || n.split(':')[0] === model.split(':')[0]);
      if (!found) return { ok: false, message: `Ollama is running, but “${model}” is not installed. In a terminal run: ollama pull ${model}${names.length ? `. Installed: ${names.join(', ')}` : ''}` };
      return { ok: true, model };
    }
    return { ok: false, message: 'Choose a suggestion provider first.' };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

module.exports = { PROVIDERS, suggestAltText, suggestText, testConnection, prepareImage, clean };
