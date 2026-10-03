'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const testTemporary = require('../scripts/local-test-env.cjs')().temp;
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { EventEmitter } = require('node:events');

// Exercise the real main-process handlers with only the Electron UI boundary
// replaced. No dialog is opened and no real profile or source file is changed.
test('desktop IPC restricts write destinations, protects originals, and keeps local navigation and close guards', async t => {
  const temporary = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-native-'));
  const root = path.resolve(__dirname, '..');
  const indexUrl = pathToFileURL(path.join(root, 'build', 'index.html')).href;
  const handlers = new Map();
  const queuedOpen = [], queuedSave = [];
  const sent = [];
  let instance, networkGuard;
  const ipcMain = new EventEmitter();
  ipcMain.handle = (name, fn) => handlers.set(name, fn);
  class BrowserWindow extends EventEmitter {
    constructor(options) {
      super(); this.options = options; instance = this;
      this.webContents = new EventEmitter();
      this.webContents.send = (...args) => sent.push(args);
      this.webContents.getURL = () => indexUrl;
      this.webContents.setWindowOpenHandler = fn => { this.popupGuard = fn; };
    }
    setTitle() {}
    setDocumentEdited() {}
    loadURL(url) { assert.equal(url, indexUrl); }
    show() {}
    isMinimized(){return false;}
    focus(){}
    close() { this.closeCalled = true; }
  }
  const app = new EventEmitter();
  const paths = { appData: temporary, userData: temporary };
  Object.assign(app, { setName() {}, setPath: (key, value) => { paths[key] = value; }, getPath: key => paths[key], commandLine: { appendSwitch() {} }, requestSingleInstanceLock: () => true, whenReady: () => Promise.resolve(), quit() {}, isPackaged: true });
  const defaultSession = new EventEmitter();
  Object.assign(defaultSession, { setPermissionRequestHandler: fn => { defaultSession.permission = fn; }, setPermissionCheckHandler: fn => { defaultSession.permissionCheck = fn; }, setDevicePermissionHandler: fn => { defaultSession.devicePermission = fn; }, webRequest: { onBeforeRequest: fn => { networkGuard = fn; } } });
  const mock = { app, BrowserWindow, ipcMain, session: { defaultSession }, Menu: { buildFromTemplate: template => template, setApplicationMenu() {} }, dialog: { showOpenDialog: async () => queuedOpen.shift(), showSaveDialog: async () => queuedSave.shift(), showErrorBox() {} } };
  const electronKey = require.resolve('electron');
  const previousModule = require.cache[electronKey];
  const previousProfile = process.env.STADIO_USER_DATA;
  try {
    process.env.STADIO_USER_DATA = temporary;
    require.cache[electronKey] = { id: electronKey, filename: electronKey, loaded: true, exports: mock };
    require('../desktop/main.cjs');
    await new Promise(resolve => setImmediate(resolve));
  } finally {
    if (previousModule) require.cache[electronKey] = previousModule; else delete require.cache[electronKey];
    if (previousProfile === undefined) delete process.env.STADIO_USER_DATA; else process.env.STADIO_USER_DATA = previousProfile;
  }
  const event = { sender: instance.webContents, senderFrame: { url: indexUrl } };
  const call = (name, ...args) => handlers.get(`stadio:${name}`)(event, ...args);
  const data = { format: 'dake-stadio', version: 1, width: 600, height: 400, name: 'ネイティブ保存', canvas: { objects: [] } };
  const raster = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';

  await t.test('unknown sender and unapproved project paths cannot access privileged file operations', async () => {
    await assert.rejects(handlers.get('stadio:getRecovery')({ sender: {}, senderFrame: { url: indexUrl } }), /許可/);
    await assert.rejects(call('saveProject', { path: path.join(temporary, 'not-selected.dake'), data }), /許可/);
  });

  await t.test('Save As establishes one authorized project and canceled saves create nothing', async () => {
    queuedSave.push({ canceled: true });
    assert.equal(await call('saveProject', { data, saveAs: true }), null);
    const destination = path.join(temporary, '作品.dake');
    queuedSave.push({ canceled: false, filePath: destination });
    assert.equal((await call('saveProject', { data, saveAs: true })).path, destination);
    await call('saveProject', { path: destination, data: { ...data, name: '次の保存' } });
    assert.equal(JSON.parse(await fs.readFile(`${destination}.bak`)).name, 'ネイティブ保存');
  });

  await t.test('imported image cannot be overwritten by native export, including a hard-link alias', async () => {
    const original = path.join(temporary, '元画像.png');
    const sourceBytes = Buffer.from(raster.split(',')[1], 'base64');
    await fs.writeFile(original, sourceBytes);
    queuedOpen.push({ canceled: false, filePaths: [original] });
    const imported = await call('importFiles');
    assert.equal(imported[0].data, raster);
    queuedSave.push({ canceled: false, filePath: original });
    await assert.rejects(call('exportFile', { format: 'png', data: raster, name: '画像' }), /元ファイル/);
    const alias = path.join(temporary, '別名.png');
    await fs.link(original, alias);
    queuedSave.push({ canceled: false, filePath: alias });
    await assert.rejects(call('exportFile', { format: 'png', data: raster, name: '画像' }), /元ファイル/);
    assert.deepEqual(await fs.readFile(original), sourceBytes);
    const exportPath = path.join(temporary, '書き出し.png');
    queuedSave.push({ canceled: false, filePath: exportPath });
    assert.equal((await call('exportFile', { format: 'png', data: raster, name: '画像' })).path, exportPath);
    assert.deepEqual(await fs.readFile(exportPath), sourceBytes);
  });

  await t.test('second-instance requests are tokenized, validated, acknowledged and grant only their project path',async()=>{
    const destination=path.join(temporary,'起動作品.dake');await fs.writeFile(destination,JSON.stringify(data));
    app.emit('second-instance',{},['app.exe',destination,'--ignore.dake'],temporary);
    const requests=await call('getOpenRequests');assert.equal(requests.length,1);assert.equal(requests[0].name,'起動作品.dake');
    await assert.rejects(call('readOpenRequest',999));
    const loaded=await call('readOpenRequest',requests[0].id);assert.equal(loaded.data.name,data.name);
    const version=await call('saveProject',{path:destination,data,version:true});assert.match(version.path,/_v001\.dake$/);
    await call('ackOpenRequest',requests[0].id);assert.deepEqual(await call('getOpenRequests'),[]);
  });

  await t.test('recovery uses the isolated local profile and explicit discard clears it', async () => {
    assert.equal(await call('autosave', data), true);
    assert.equal((await call('getRecovery')).data.name, data.name);
    assert.ok(await fs.stat(path.join(temporary, 'recovery.json')));
    await call('discardRecovery');
    assert.equal(await call('getRecovery'), null);
  });

  await t.test('network, foreign local files, permissions, and popups are blocked', () => {
    for (const url of ['https://example.com/', 'http://localhost:9000/', 'file:///C:/Windows/win.ini']) networkGuard({ url }, result => assert.equal(result.cancel, true));
    networkGuard({ url: indexUrl }, result => assert.equal(result.cancel, false));
    assert.deepEqual(instance.popupGuard(), { action: 'deny' });
    assert.equal(defaultSession.permissionCheck(), false);
    assert.equal(defaultSession.devicePermission(), false);
    assert.equal(instance.options.webPreferences.sandbox, true);
    assert.equal(instance.options.webPreferences.contextIsolation, true);
    assert.equal(instance.options.webPreferences.nodeIntegration, false);
  });

  await t.test('dropped image sources stay protected; project drops enable save and mixed drops are rejected', async () => {
    const original=path.join(temporary,'drop.png');await fs.writeFile(original,Buffer.from(raster.split(',')[1],'base64'));
    assert.equal((await call('importDroppedFiles',[original]))[0].kind,'image');
    queuedSave.push({canceled:false,filePath:original});await assert.rejects(call('exportFile',{format:'png',data:raster}),/元ファイル/);
    const project=path.join(temporary,'drop.dake');await fs.writeFile(project,JSON.stringify(data));
    assert.equal((await call('importDroppedFiles',[project])).kind,'project');
    await call('saveProject',{path:project,data:{...data,name:'drop saved'}});
    assert.equal(JSON.parse(await fs.readFile(project)).name,'drop saved');
    await assert.rejects(call('importDroppedFiles',[project,original]));
    await assert.rejects(call('importDroppedFiles',new Array(33).fill(original)));
    await assert.rejects(call('importDroppedFiles',['https://example.org/a.png']));
  });
  await t.test('local font permission is granted only to the trusted editor document', () => {
    assert.equal(defaultSession.permissionCheck(instance.webContents,'local-fonts'),true);
    assert.equal(defaultSession.permissionCheck({getURL:()=>indexUrl},'local-fonts'),false);
    assert.equal(defaultSession.permissionCheck(instance.webContents,'camera'),false);
    instance.webContents.getURL=()=> 'file:///C:/other.html';assert.equal(defaultSession.permissionCheck(instance.webContents,'local-fonts'),false);
    instance.webContents.getURL=()=>indexUrl;
  });

  await t.test('native close waits for renderer cleanup even for a clean document', () => {
    let prevented = false;
    instance.emit('close', { preventDefault: () => { prevented = true; } });
    assert.equal(prevented, true);
    assert.equal(sent.at(-1)[0], 'stadio:closeRequest');
    ipcMain.emit('stadio:closeReady', event);
    assert.equal(instance.closeCalled, true);
  });
});
