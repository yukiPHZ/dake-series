'use strict';
const { contextBridge, ipcRenderer, webUtils } = require('electron');
let localFontPromise;
const localFonts = new Map();
function subscribe(channel, listener) {
  if (typeof listener !== 'function') return () => {};
  const handler = (_event, payload) => listener(payload);
  ipcRenderer.on(channel, handler);
  return () => ipcRenderer.removeListener(channel, handler);
}
function fontDescriptor(font) {
  const nativeStyle = font.style || 'Regular', lower = nativeStyle.toLowerCase();
  let weight = 400;
  if (/thin/.test(lower)) weight = 100;
  else if (/extra.?light|ultra.?light/.test(lower)) weight = 200;
  else if (/light/.test(lower)) weight = 300;
  else if (/medium/.test(lower)) weight = 500;
  else if (/semi.?bold|demi.?bold/.test(lower)) weight = 600;
  else if (/extra.?bold|ultra.?bold/.test(lower)) weight = 800;
  else if (/black|heavy/.test(lower)) weight = 900;
  else if (/bold/.test(lower)) weight = 700;
  return { id: 'local:' + font.postscriptName, source: 'local', family: font.family, fullName: font.fullName, postscriptName: font.postscriptName, nativeStyle, style: /italic|oblique/.test(lower) ? 'italic' : 'normal', weight };
}
async function listLocalFonts() {
  if (typeof window.queryLocalFonts !== 'function') throw new Error('FONT_ACCESS_UNAVAILABLE');
  if (!localFontPromise) localFontPromise = window.queryLocalFonts().then(fonts => {
    localFonts.clear();
    for (const font of fonts) localFonts.set('local:' + font.postscriptName, font);
    return fonts.map(fontDescriptor);
  }).catch(error => { localFontPromise = null; throw error; });
  return localFontPromise;
}
async function getLocalFont(id) {
  await listLocalFonts();
  const font = localFonts.get(String(id));
  if (!font) throw new Error('FONT_UNAVAILABLE');
  const blob = await font.blob();
  if (blob.size > 64 * 1024 * 1024) throw new Error('FONT_INVALID');
  return ipcRenderer.invoke('stadio:prepareLocalFont', { descriptor: fontDescriptor(font), bytes: await blob.arrayBuffer() });
}
contextBridge.exposeInMainWorld('stadio', Object.freeze({
  getPendingRecoveries:()=>ipcRenderer.invoke('stadio:getPendingRecoveries'),
  readPendingRecovery:id=>ipcRenderer.invoke('stadio:readPendingRecovery',id),
  discardPendingRecovery:id=>ipcRenderer.invoke('stadio:discardPendingRecovery',id),
  prepareExport: request=>ipcRenderer.invoke('stadio:prepareExport',request),
  savePreparedExport: token=>ipcRenderer.invoke('stadio:savePreparedExport',token),
  releasePreparedExport: token=>ipcRenderer.invoke('stadio:releasePreparedExport',token),
  importComponent: mode=>ipcRenderer.invoke('stadio:importComponent',mode),
  refreshComponent: token=>ipcRenderer.invoke('stadio:refreshComponent',token),
  autosaveWorkspace: data=>ipcRenderer.invoke('stadio:autosaveWorkspace',data),
  getWorkspaceRecovery: ()=>ipcRenderer.invoke('stadio:getWorkspaceRecovery'),
  discardWorkspaceRecovery: ()=>ipcRenderer.invoke('stadio:discardWorkspaceRecovery'),
  getOpenRequests: () => ipcRenderer.invoke('stadio:getOpenRequests'),
  readOpenRequest: id => ipcRenderer.invoke('stadio:readOpenRequest', id),
  ackOpenRequest: id => ipcRenderer.invoke('stadio:ackOpenRequest', id),
  onOpenRequest: listener => subscribe('stadio:openRequest', listener),
  listCachedFonts: () => ipcRenderer.invoke('stadio:listCachedFonts'),
  getCachedFont: key => ipcRenderer.invoke('stadio:getCachedFont', key),
  openProject: () => ipcRenderer.invoke('stadio:openProject'),
  saveProject: (request) => ipcRenderer.invoke('stadio:saveProject', request),
  importFiles: () => ipcRenderer.invoke('stadio:importFiles'),
  importDroppedFiles: files => ipcRenderer.invoke('stadio:importDroppedFiles', Array.from(files || []).map(file => webUtils.getPathForFile(file)).filter(Boolean)),
  exportFile: (request) => ipcRenderer.invoke('stadio:exportFile', request),
  printArtwork: request => ipcRenderer.invoke('stadio:printArtwork', request),
  listLocalFonts, getLocalFont,
  listGoogleFonts: () => ipcRenderer.invoke('stadio:listGoogleFonts'),
  downloadGoogleFont: request => ipcRenderer.invoke('stadio:downloadGoogleFont', request),
  cancelFontDownload: id => ipcRenderer.invoke('stadio:cancelFontDownload', id),
  onFontProgress: listener => subscribe('stadio:fontProgress', listener),
  autosave: (data) => ipcRenderer.invoke('stadio:autosave', data),
  getRecovery: () => ipcRenderer.invoke('stadio:getRecovery'),
  discardRecovery: () => ipcRenderer.invoke('stadio:discardRecovery'),
  setDirty: (dirty) => ipcRenderer.send('stadio:setDirty', !!dirty),
  onMenu: (listener) => subscribe('stadio:menu', listener),
  onCloseRequest: (listener) => subscribe('stadio:closeRequest', listener),
  closeReady: () => ipcRenderer.send('stadio:closeReady')
}));

