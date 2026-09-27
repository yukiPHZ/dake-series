'use strict';
const { BrowserWindow } = require('electron');
const { Worker, isMarkedAsUntransferable } = require('node:worker_threads');
const path = require('node:path');
const { encodeExport } = require('./storage.cjs');
const UI_TEXT = require('../src/ui-text.json').desktop;
function dimensions(request) {
  const width = Number(request?.width), height = Number(request?.height), dpi = Number(request?.dpi || 96);
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || width > 8192 || height > 8192 || width * height > 32000000 || !Number.isFinite(dpi) || dpi < 30 || dpi > 2400) throw new Error(UI_TEXT.errorPrintSize);
  const widthMm = width / dpi * 25.4, heightMm = height / dpi * 25.4;
  if (widthMm < .36 || heightMm < .36 || widthMm > 2000 || heightMm > 2000) throw new Error(UI_TEXT.errorPrintSize);
  const png = encodeExport('png', request.data);
  if (png.length < 24 || png.readUInt32BE(16) !== width || png.readUInt32BE(20) !== height) throw new Error(UI_TEXT.errorPrintSize);
  return { width, height, dpi, widthMm, heightMm, png };
}
async function printWindow(request) {
  const size = dimensions(request);
  const window = new BrowserWindow({ show: false, width: 800, height: 600, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, webSecurity: true, spellcheck: false } });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  const html = '<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src \'none\';img-src data:;style-src \'unsafe-inline\';base-uri \'none\'"><style>@page{size:' + size.widthMm + 'mm ' + size.heightMm + 'mm;margin:0}html,body{margin:0;padding:0;width:' + size.widthMm + 'mm;height:' + size.heightMm + 'mm;overflow:hidden;background:#fff}img{display:block;width:100%;height:100%}</style></head><body><img src="' + request.data + '"></body></html>';
  try {
    await window.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
    await window.webContents.executeJavaScript('Promise.all(Array.from(document.images).map(image=>image.decode()))');
    return { window, size };
  } catch (error) { window.destroy(); throw error; }
}
async function toPDF(request) {
  const size = dimensions(request), bytes = size.png;
  return new Promise((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, 'print-pdf-worker.cjs'), { workerData: { bytes, width: size.width, height: size.height, dpi: size.dpi }, transferList: isMarkedAsUntransferable(bytes.buffer) ? [] : [bytes.buffer] });
    let completed = false;
    worker.once('message', result => { completed = true; if (result.error) reject(new Error(UI_TEXT.errorPrint)); else resolve(Buffer.from(result.bytes)); });
    worker.once('error', () => { completed = true; reject(new Error(UI_TEXT.errorPrint)); });
    worker.once('exit', () => { if (!completed) reject(new Error(UI_TEXT.errorPrint)); });
  });
}
async function printArtwork(request) {
  const { window, size } = await printWindow(request);
  try {
    return await new Promise((resolve, reject) => window.webContents.print({ silent: false, printBackground: true, margins: { marginType: 'none' }, pageSize: { width: Math.round(size.widthMm * 1000), height: Math.round(size.heightMm * 1000) } }, (success, reason) => {
      if (success) resolve({ printed: true, cancelled: false });
      else if (/cancel/i.test(reason || '')) resolve({ printed: false, cancelled: true });
      else reject(new Error(UI_TEXT.errorPrint));
    }));
  } finally { window.destroy(); }
}
module.exports = { dimensions, toPDF, printArtwork };

