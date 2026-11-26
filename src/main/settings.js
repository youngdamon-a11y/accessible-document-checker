'use strict';
// Local settings. The Claude API key is encrypted with the operating system's
// keychain (Electron safeStorage) and never leaves this computer except when it
// is sent to the Anthropic API to request a suggestion.
const fs = require('fs');
const path = require('path');
const { app, safeStorage } = require('electron');

const DEFAULTS = {
  defaultLanguage: 'en-GB',
  model: 'claude-opus-5-5',
  outputFolder: '', // '' = same folder as the original
  aiEnabled: true,
};
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
  const { apiKeyEnc, apiKeyPlain, ...rest } = raw;
  return { ...DEFAULTS, ...rest, hasApiKey: !!(apiKeyEnc || apiKeyPlain), keyStorage: safeStorage.isEncryptionAvailable() ? 'encrypted' : 'plain' };
}
function set(patch) {
  const raw = readRaw();
  for (const k of Object.keys(DEFAULTS)) if (patch[k] !== undefined) raw[k] = patch[k];
  writeRaw(raw);
  return get();
}
function setApiKey(key) {
  const raw = readRaw();
  delete raw.apiKeyEnc; delete raw.apiKeyPlain;
  if (key) {
    if (safeStorage.isEncryptionAvailable()) raw.apiKeyEnc = safeStorage.encryptString(key).toString('base64');
    else raw.apiKeyPlain = key; // no OS keychain available (some Linux desktops)
  }
  writeRaw(raw);
  return get();
}
function getApiKey() {
  const raw = readRaw();
  if (raw.apiKeyEnc) {
    try { return safeStorage.decryptString(Buffer.from(raw.apiKeyEnc, 'base64')); } catch { return null; }
  }
  return raw.apiKeyPlain || process.env.ANTHROPIC_API_KEY || null;
}

module.exports = { get, set, setApiKey, getApiKey, DEFAULTS };
