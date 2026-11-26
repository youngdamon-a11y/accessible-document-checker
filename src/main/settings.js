'use strict';
// Local settings. API keys (Claude, OpenAI) are encrypted with the operating system's
// keychain (Electron safeStorage) and never leave this computer except when sent to
// that provider to request a suggestion. The local Ollama option needs no key.
const fs = require('fs');
const path = require('path');
const { app, safeStorage } = require('electron');

const DEFAULTS = {
  defaultLanguage: 'en-GB',
  provider: 'anthropic', // 'anthropic' | 'openai' | 'ollama' | 'off'
  model: 'claude-opus-5-5', // Claude model
  openaiModel: 'gpt-5',
  ollamaUrl: 'http://localhost:11434',
  ollamaModel: 'gemma3',
  outputFolder: '', // '' = same folder as the original
  aiEnabled: true, // kept for older settings files; provider 'off' turns suggestions off
};
// Where each provider's key is stored in settings.json.
const KEY_FIELDS = { anthropic: ['apiKeyEnc', 'apiKeyPlain', 'ANTHROPIC_API_KEY'], openai: ['openaiKeyEnc', 'openaiKeyPlain', 'OPENAI_API_KEY'] };
const file = () => path.join(app.getPath('userData'), 'settings.json');

function readRaw() {
  try { return JSON.parse(fs.readFileSync(file(), 'utf8')); } catch { return {}; }
}
function writeRaw(obj) {
  fs.mkdirSync(path.dirname(file()), { recursive: true });
  fs.writeFileSync(file(), JSON.stringify(obj, null, 2), { mode: 0o600 });
}
function get() {
  const raw = readRaw();
  const out = { ...DEFAULTS };
  for (const k of Object.keys(DEFAULTS)) if (raw[k] !== undefined) out[k] = raw[k];
  if (raw.aiEnabled === false && raw.provider === undefined) out.provider = 'off';
  out.hasApiKey = !!(raw.apiKeyEnc || raw.apiKeyPlain);
  out.hasOpenaiKey = !!(raw.openaiKeyEnc || raw.openaiKeyPlain);
  out.keyStorage = safeStorage.isEncryptionAvailable() ? 'encrypted' : 'plain';
  return out;
}
function set(patch) {
  const raw = readRaw();
  for (const k of Object.keys(DEFAULTS)) if (patch[k] !== undefined) raw[k] = patch[k];
  if (patch.provider !== undefined) raw.aiEnabled = patch.provider !== 'off';
  writeRaw(raw);
  return get();
}
function setApiKey(key, provider = 'anthropic') {
  const [enc, plain] = KEY_FIELDS[provider] || KEY_FIELDS.anthropic;
  const raw = readRaw();
  delete raw[enc]; delete raw[plain];
  if (key) {
    if (safeStorage.isEncryptionAvailable()) raw[enc] = safeStorage.encryptString(key).toString('base64');
    else raw[plain] = key; // no OS keychain available (some Linux desktops)
  }
  writeRaw(raw);
  return get();
}
function getApiKey(provider = 'anthropic') {
  const [enc, plain, env] = KEY_FIELDS[provider] || KEY_FIELDS.anthropic;
  const raw = readRaw();
  if (raw[enc]) {
    try { return safeStorage.decryptString(Buffer.from(raw[enc], 'base64')); } catch { return null; }
  }
  return raw[plain] || process.env[env] || null;
}
/** Everything the AI module needs for the chosen provider. */
function aiConfig(override = {}) {
  const s = { ...get(), ...override };
  const provider = s.provider;
  if (provider === 'anthropic') return { provider, model: s.model, apiKey: override.apiKey || getApiKey('anthropic') };
  if (provider === 'openai') return { provider, model: s.openaiModel, apiKey: override.apiKey || getApiKey('openai') };
  if (provider === 'ollama') return { provider, model: s.ollamaModel, url: s.ollamaUrl };
  return { provider: 'off' };
}

module.exports = { get, set, setApiKey, getApiKey, aiConfig, DEFAULTS };
