'use strict';
const { contextBridge, ipcRenderer } = require('electron');

function subscribe(channel, listener) {
  if (typeof listener !== 'function') return () => {};
  const handler = (_event, payload) => listener(payload);
  ipcRenderer.on(channel, handler);
  return () => ipcRenderer.removeListener(channel, handler);
}

contextBridge.exposeInMainWorld('stadio', Object.freeze({
  openProject: () => ipcRenderer.invoke('stadio:openProject'),
  saveProject: (request) => ipcRenderer.invoke('stadio:saveProject', request),
  importFiles: () => ipcRenderer.invoke('stadio:importFiles'),
  exportFile: (request) => ipcRenderer.invoke('stadio:exportFile', request),
  autosave: (data) => ipcRenderer.invoke('stadio:autosave', data),
  getRecovery: () => ipcRenderer.invoke('stadio:getRecovery'),
  discardRecovery: () => ipcRenderer.invoke('stadio:discardRecovery'),
  setDirty: (dirty) => ipcRenderer.send('stadio:setDirty', !!dirty),
  onMenu: (listener) => subscribe('stadio:menu', listener),
  onCloseRequest: (listener) => subscribe('stadio:closeRequest', listener),
  closeReady: () => ipcRenderer.send('stadio:closeReady')
}));
