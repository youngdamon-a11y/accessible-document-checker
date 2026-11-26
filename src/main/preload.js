'use strict';
const { contextBridge, ipcRenderer, webUtils } = require('electron');

const call = async (channel, ...args) => {
  const r = await ipcRenderer.invoke(channel, ...args);
  if (!r.ok) throw new Error(r.error);
  return r.value;
};

contextBridge.exposeInMainWorld('api', {
  pickFiles: () => call('docs:pick'),
  openFiles: (paths) => call('docs:open', paths),
  pathForFile: (file) => webUtils.getPathForFile(file),
  closeDoc: (id) => call('doc:close', id),
  recheck: (id) => call('doc:recheck', id),
  fix: (id, issueId, input) => call('doc:fix', id, issueId, input),
  fixAll: (id) => call('doc:fixAll', id),
  undo: (id) => call('doc:undo', id),
  humanCheck: (id, checkId, status, note) => call('doc:humanCheck', id, checkId, status, note),
  preview: (id, issueId) => call('doc:preview', id, issueId),
  suggest: (id, issueId) => call('ai:suggest', id, issueId),
  save: (id) => call('doc:save', id),
  saveAs: (id) => call('doc:saveAs', id),
  exportReport: (id, kind) => call('report:export', id, kind),
  getSettings: () => call('settings:get'),
  setSettings: (patch) => call('settings:set', patch),
  setApiKey: (key) => call('settings:setKey', key),
  testApiKey: (key) => call('settings:testKey', key),
  pickFolder: () => call('settings:pickFolder'),
  showInFolder: (p) => call('shell:show', p),
  openFile: (p) => call('shell:open', p),
  openExternal: (url) => call('shell:external', url),
  onMenu: (fn) => ipcRenderer.on('menu', (_e, what) => fn(what)),
  onProgress: (fn) => ipcRenderer.on('progress', (_e, p) => fn(p)),
});
