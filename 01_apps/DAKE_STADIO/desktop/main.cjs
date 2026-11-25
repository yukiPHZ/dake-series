'use strict';

const { app, BrowserWindow, Menu, dialog, ipcMain, session } = require('electron');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const { pathToFileURL, fileURLToPath } = require('node:url');
const UI_TEXT = require('../src/ui-text.json').desktop;
const storage = require('./storage.cjs');

app.setName(UI_TEXT.appName);
if (process.env.STADIO_USER_DATA) app.setPath('userData', path.resolve(process.env.STADIO_USER_DATA));
else app.setPath('userData', path.join(app.getPath('appData'), UI_TEXT.appName));
fsSync.mkdirSync(app.getPath('userData'), { recursive: true });
app.commandLine.appendSwitch('disable-component-update');
app.commandLine.appendSwitch('disable-background-networking');
app.commandLine.appendSwitch('lang', 'ja');

const appRoot = path.resolve(__dirname, '..');
const buildRoot = path.join(appRoot, 'build');
const indexUrl = pathToFileURL(path.join(buildRoot, 'index.html')).href;
const recoveryPath = path.join(app.getPath('userData'), 'recovery.json');
const writableProjects = new Set();
const importedSources = new Set();
const importedFileIds = new Set();
let window;
let dirty = false;
let closeAllowed = false;
let rendererGone = false;
let writeQueue = Promise.resolve();

function enqueueWrite(task) {
  const result = writeQueue.then(task, task);
  writeQueue = result.catch(() => {});
  return result;
}

function samePathKey(filePath) { return path.resolve(filePath).toLowerCase(); }
async function canonicalKey(filePath) {
  try { return samePathKey(await fs.realpath(filePath)); }
  catch (error) { if (error.code === 'ENOENT') return samePathKey(filePath); throw error; }
}

function localizedError(key) { return new Error(UI_TEXT[key]); }
async function rememberSource(filePath) {
  importedSources.add(await canonicalKey(filePath));
  const stat = await fs.stat(filePath);
  importedFileIds.add(`${stat.dev}:${stat.ino}`);
}

async function protectSource(filePath) {
  if (importedSources.has(await canonicalKey(filePath))) throw localizedError('errorSource');
  try {
    const stat = await fs.stat(filePath);
    if (importedFileIds.has(`${stat.dev}:${stat.ino}`)) throw localizedError('errorSource');
  } catch (error) { if (error.code !== 'ENOENT') throw error; }
}

function ensureExtension(filePath, extensions) {
  if (!extensions.includes(path.extname(filePath).slice(1).toLowerCase())) throw localizedError('errorExtension');
}

function senderAllowed(event) {
  return window && event.sender === window.webContents && event.senderFrame && event.senderFrame.url === indexUrl;
}

function handle(channel, callback) {
  ipcMain.handle(`stadio:${channel}`, async (event, ...args) => {
    if (!senderAllowed(event)) throw localizedError('errorPath');
    return callback(...args);
  });
}

function registerHandlers() {
  handle('openProject', async () => {
    const selected = await dialog.showOpenDialog(window, {
      title: UI_TEXT.dialogOpen, properties: ['openFile'],
      filters: [{ name: UI_TEXT.filterProject, extensions: ['dake', 'bak'] }]
    });
    if (selected.canceled || !selected.filePaths[0]) return null;
    const filePath = selected.filePaths[0];
    const data = await storage.readProject(filePath);
    // A .bak is readable but never the automatic overwrite destination.
    if (path.extname(filePath).toLowerCase() === '.dake') writableProjects.add(await canonicalKey(filePath));
    return { path: filePath, data };
  });

  handle('saveProject', (request) => enqueueWrite(async () => {
    if (!request || typeof request !== 'object') throw localizedError('errorProject');
    storage.validateProject(request.data);
    let filePath = typeof request.path === 'string' ? request.path : null;
    if (!filePath || request.saveAs || path.extname(filePath).toLowerCase() !== '.dake') {
      const selected = await dialog.showSaveDialog(window, {
        title: UI_TEXT.dialogSave,
        defaultPath: filePath && path.extname(filePath).toLowerCase() === '.dake' ? filePath : `${storage.safeName(request.data.name)}.dake`,
        filters: [{ name: UI_TEXT.filterProject, extensions: ['dake'] }],
        properties: ['showOverwriteConfirmation', 'createDirectory']
      });
      if (selected.canceled || !selected.filePath) return null;
      filePath = selected.filePath;
      ensureExtension(filePath, ['dake']);
      await protectSource(filePath);
      writableProjects.add(await canonicalKey(filePath));
    } else if (!writableProjects.has(await canonicalKey(filePath))) throw localizedError('errorPath');
    await protectSource(filePath);
    // A backup path can also have been imported, and must receive equal protection.
    await protectSource(`${filePath}.bak`);
    await storage.writeProject(filePath, request.data);
    writableProjects.add(await canonicalKey(filePath));
    return { path: filePath };
  }));

  handle('importFiles', async () => {
    const selected = await dialog.showOpenDialog(window, {
      title: UI_TEXT.dialogImport, properties: ['openFile', 'multiSelections'],
      filters: [{ name: UI_TEXT.filterImport, extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg'] }]
    });
    if (selected.canceled) return [];
    if (selected.filePaths.length > 32) throw localizedError('errorLarge');
    const files = [];
    let totalBytes = 0;
    for (const filePath of selected.filePaths) {
      const bytes = await storage.readLimited(filePath);
      totalBytes += bytes.length;
      if (totalBytes > storage.MAX_BYTES) throw localizedError('errorLarge');
      const extension = path.extname(filePath).toLowerCase();
      if (extension === '.svg') {
        files.push({ name: path.basename(filePath), kind: 'svg', data: storage.validateSvg(bytes.toString('utf8')) });
      } else {
        const mime = { '.png': 'png', '.jpg': 'jpeg', '.jpeg': 'jpeg', '.webp': 'webp', '.gif': 'gif', '.bmp': 'bmp' }[extension];
        if (!mime) throw localizedError('errorImage');
        files.push({ name: path.basename(filePath), kind: 'image', data: `data:image/${mime};base64,${bytes.toString('base64')}` });
      }
      await rememberSource(filePath);
    }
    return files;
  });

  handle('exportFile', (request) => enqueueWrite(async () => {
    if (!request || typeof request !== 'object') throw localizedError('errorExport');
    const formats = { png: ['png'], jpeg: ['jpg', 'jpeg'], jpg: ['jpg', 'jpeg'], webp: ['webp'], svg: ['svg'] };
    const extensions = formats[request.format];
    if (!extensions) throw localizedError('errorExport');
    const bytes = storage.encodeExport(request.format, request.data);
    const selected = await dialog.showSaveDialog(window, {
      title: UI_TEXT.dialogExport,
      defaultPath: `${storage.safeName(request.name)}.${extensions[0]}`,
      filters: [{ name: request.format.toUpperCase(), extensions }],
      properties: ['showOverwriteConfirmation', 'createDirectory']
    });
    if (selected.canceled || !selected.filePath) return null;
    ensureExtension(selected.filePath, extensions);
    await protectSource(selected.filePath);
    await storage.atomicWrite(selected.filePath, bytes);
    return { path: selected.filePath };
  }));

  handle('autosave', (data) => enqueueWrite(async () => {
    await storage.writeRecovery(recoveryPath, data);
    return true;
  }));

  handle('getRecovery', () => enqueueWrite(() => storage.readRecovery(recoveryPath)));

  handle('discardRecovery', () => enqueueWrite(async () => {
    await storage.discardRecovery(recoveryPath);
    return true;
  }));

  ipcMain.on('stadio:setDirty', (event, value) => {
    if (!senderAllowed(event)) return;
    dirty = !!value;
    window.setDocumentEdited(dirty);
    window.setTitle(`${dirty ? '● ' : ''}${UI_TEXT.appName}`);
  });
  ipcMain.on('stadio:closeReady', (event) => {
    if (!senderAllowed(event)) return;
    closeAllowed = true;
    window.close();
  });
}

function createMenu() {
  const action = (label, accelerator, command) => ({ label, accelerator, click: () => window?.webContents.send('stadio:menu', command) });
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    { label: UI_TEXT.menuFile, submenu: [
      action(UI_TEXT.menuNew, 'CmdOrCtrl+N', 'new'),
      action(UI_TEXT.menuOpen, 'CmdOrCtrl+O', 'open'),
      action(UI_TEXT.menuSave, 'CmdOrCtrl+S', 'save'),
      action(UI_TEXT.menuSaveAs, 'CmdOrCtrl+Shift+S', 'saveAs'),
      { type: 'separator' },
      action(UI_TEXT.menuImport, 'CmdOrCtrl+I', 'import'),
      action(UI_TEXT.menuExport, 'CmdOrCtrl+Shift+E', 'export'),
      { type: 'separator' },
      { label: UI_TEXT.menuClose, accelerator: 'Alt+F4', click: () => window?.close() }
    ] },
    { label: UI_TEXT.menuEdit, submenu: [
      action(UI_TEXT.menuUndo, 'CmdOrCtrl+Z', 'undo'),
      action(UI_TEXT.menuRedo, 'CmdOrCtrl+Shift+Z', 'redo')
    ] },
    { label: UI_TEXT.menuView, submenu: [
      action(UI_TEXT.menuFit, 'CmdOrCtrl+0', 'fit'),
      action(UI_TEXT.menuActual, 'CmdOrCtrl+1', 'actual'),
      { label: UI_TEXT.menuFullscreen, role: 'togglefullscreen' },
      ...(!app.isPackaged ? [{ label: UI_TEXT.menuDevTools, role: 'toggleDevTools' }] : [])
    ] }
  ]));
}

function createWindow() {
  window = new BrowserWindow({
    width: 1440, height: 980, minWidth: 1060, minHeight: 720,
    title: UI_TEXT.appName, backgroundColor: '#eef1f5',
    icon: path.join(appRoot, 'assets', 'dake_icon.ico'),
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false, contextIsolation: true, sandbox: true,
      webSecurity: true, allowRunningInsecureContent: false,
      devTools: !app.isPackaged, spellcheck: false
    }
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event, target) => { if (target !== indexUrl) event.preventDefault(); });
  window.webContents.on('will-attach-webview', event => event.preventDefault());
  window.webContents.on('render-process-gone', () => {
    rendererGone = true;
    dialog.showErrorBox(UI_TEXT.errorTitle, UI_TEXT.errorRenderer);
  });
  window.once('ready-to-show', () => window.show());
  window.on('close', event => {
    if (!closeAllowed && !rendererGone) {
      event.preventDefault();
      window.webContents.send('stadio:closeRequest');
    }
  });
  window.on('closed', () => { window = null; });
  createMenu();
  window.loadURL(indexUrl);
}

if (!app.requestSingleInstanceLock()) app.quit();
else {
  app.on('second-instance', () => { if (window) { if (window.isMinimized()) window.restore(); window.focus(); } });
  app.whenReady().then(() => {
    session.defaultSession.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
    session.defaultSession.setPermissionCheckHandler(() => false);
    session.defaultSession.setDevicePermissionHandler(() => false);
    session.defaultSession.webRequest.onBeforeRequest((details, callback) => {
      let allowed = details.url.startsWith('data:') || details.url.startsWith('blob:') || details.url === 'about:blank' || details.url.startsWith('devtools:');
      if (details.url.startsWith('file:')) {
        try {
          const filePath = path.resolve(fileURLToPath(details.url));
          allowed = filePath.startsWith(`${buildRoot}${path.sep}`);
        } catch { allowed = false; }
      }
      callback({ cancel: !allowed });
    });
    session.defaultSession.on('will-download', event => event.preventDefault());
    registerHandlers();
    createWindow();
  }).catch(error => { dialog.showErrorBox(UI_TEXT.errorTitle, error.message || UI_TEXT.errorGeneric); app.quit(); });
  app.on('window-all-closed', () => app.quit());
}
