'use strict';

const { app, BrowserWindow, Menu, dialog, ipcMain, session } = require('electron');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const { pathToFileURL, fileURLToPath } = require('node:url');
const UI_TEXT = require('../src/ui-text.json').desktop;
const storage = require('./storage.cjs');
const {createComponentService}=require('./components.cjs');
const workspaceStorage=require('./workspace-storage.cjs');
const {createRecoveryManager}=require('./recovery-manager.cjs');
const { createFontService } = require('./fonts.cjs');
const print = require('./print-artwork.cjs');

app.setName(UI_TEXT.appName);
if (process.env.STADIO_USER_DATA) app.setPath('userData', path.resolve(process.env.STADIO_USER_DATA));
else if (fsSync.existsSync(path.join(path.dirname(process.execPath), '.stadio-portable'))) app.setPath('userData', path.join(path.dirname(process.execPath), 'user-data'));
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
let nextOpenId = 0;
const openRequests = [];
function queueOpen(argv, cwd = process.cwd()) {
  for (const arg of argv) if (typeof arg === 'string' && !arg.startsWith('-') && /\.dake$/i.test(arg)) {
    const filePath = path.resolve(cwd, arg);
    if (!openRequests.some(r => r.path === filePath)) openRequests.push({id: ++nextOpenId, path: filePath});
  }
  window?.webContents.send('stadio:openRequest');
}
queueOpen(process.argv.slice(1));
const fontService = createFontService({ userData: app.getPath('userData'), catalogPath: path.join(appRoot, 'assets', 'google-fonts-catalog.json'), progress: payload => window?.webContents.send('stadio:fontProgress', payload) });

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

async function readImports(filePaths, allowProject = false) {
  if (!Array.isArray(filePaths) || filePaths.length > 32 || filePaths.some(value => typeof value !== 'string' || value.length > 8192)) throw localizedError('errorLarge');
  const projects = filePaths.filter(filePath => /\.(?:dake|bak)$/i.test(filePath));
  if (projects.length) {
    if (!allowProject || filePaths.length !== 1) throw localizedError('errorProject');
    const filePath = projects[0], data = await storage.readProject(filePath);
    if (path.extname(filePath).toLowerCase() === '.dake') writableProjects.add(await canonicalKey(filePath));
    return { kind: 'project', path: filePath, data };
  }
  const files = []; let totalBytes = 0;
  for (const filePath of filePaths) {
    const bytes = await storage.readLimited(filePath); totalBytes += bytes.length;
    if (totalBytes > storage.MAX_BYTES) throw localizedError('errorLarge');
    const extension = path.extname(filePath).toLowerCase();
    if (extension === '.svg') files.push({ name: path.basename(filePath), kind: 'svg', data: storage.validateSvg(bytes.toString('utf8')) });
    else {
      const mime = { '.png': 'png', '.jpg': 'jpeg', '.jpeg': 'jpeg', '.webp': 'webp', '.gif': 'gif', '.bmp': 'bmp' }[extension];
      if (!mime) throw localizedError('errorImage');
      files.push({ name: path.basename(filePath), kind: 'image', data: 'data:image/' + mime + ';base64,' + bytes.toString('base64') });
    }
    await rememberSource(filePath);
  }
  return files;
}
function registerHandlers() {
  handle('getOpenRequests', () => openRequests.map(r => ({id:r.id, name:path.basename(r.path)})));
  handle('readOpenRequest', async id => {
    const request = openRequests.find(r => r.id === id);
    if (!request) throw localizedError('errorPath');
    const data = await storage.readProject(request.path);
    writableProjects.add(await canonicalKey(request.path));
    return {path:request.path, data};
  });
  handle('ackOpenRequest', id => { const index=openRequests.findIndex(r=>r.id===id); if(index>=0)openRequests.splice(index,1); });
  handle('listCachedFonts', () => fontService.listCached());
  handle('getCachedFont', key => fontService.getCached(key));
  handle('listGoogleFonts', () => fontService.list());
  handle('downloadGoogleFont', request => fontService.download(request));
  handle('cancelFontDownload', id => fontService.cancel(id));
  handle('prepareLocalFont', request => fontService.prepareLocal(request));
  handle('importDroppedFiles', paths => readImports(paths, true));
  handle('printArtwork', request => print.printArtwork(request));
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
    // Only a separately opened normal document grants this explicit overwrite. Component import never grants it.
    if(!writableProjects.has(await canonicalKey(filePath)))await protectSource(filePath);
    // A backup path can also have been imported, and must receive equal protection.
    await protectSource(`${filePath}.bak`);
    if (request.version && request.path && !request.saveAs) filePath = await storage.writeNumberedProject(filePath, request.data);
    else await storage.writeProject(filePath, request.data);
    writableProjects.add(await canonicalKey(filePath));
    return { path: filePath };
  }));

  handle('importFiles', async () => {
    const selected = await dialog.showOpenDialog(window, {
      title: UI_TEXT.dialogImport, properties: ['openFile', 'multiSelections'],
      filters: [{ name: UI_TEXT.filterImport, extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg'] }]
    });
    return selected.canceled ? [] : readImports(selected.filePaths);
  });

  handle('exportFile', (request) => enqueueWrite(async () => {
    if (!request || typeof request !== 'object') throw localizedError('errorExport');
    const formats = { png: ['png'], jpeg: ['jpg', 'jpeg'], jpg: ['jpg', 'jpeg'], webp: ['webp'], svg: ['svg'], pdf: ['pdf'] };
    const extensions = formats[request.format];
    if (!extensions) throw localizedError('errorExport');
    const bytes = request.format === 'pdf' ? await print.toPDF(request) : storage.encodeExport(request.format, request.data);
    const selected = await dialog.showSaveDialog(window, {
      title: UI_TEXT.dialogExport,
      defaultPath: `${storage.safeName(request.name)}.${extensions[0]}`,
      filters: [{ name: request.format === 'pdf' ? UI_TEXT.filterPdf : request.format.toUpperCase(), extensions }],
      properties: ['showOverwriteConfirmation', 'createDirectory']
    });
    if (selected.canceled || !selected.filePath) return null;
    ensureExtension(selected.filePath, extensions);
    await protectSource(selected.filePath);
    await storage.atomicWrite(selected.filePath, bytes);
    return { path: selected.filePath };
  }));

  const componentService=createComponentService({readProject:storage.readProject,rememberSource});
  handle('importComponent',async(mode)=>{
    if(!['embedded','linked'].includes(mode))throw new Error(UI_TEXT.errorProject);
    const result=await dialog.showOpenDialog(window,{title:UI_TEXT.componentImport,properties:['openFile'],filters:[{name:UI_TEXT.filterProject,extensions:['dake']}]});
    if(result.canceled||!result.filePaths.length)return null;
    return componentService.importFile(result.filePaths[0],mode);
  });
  handle('refreshComponent',token=>componentService.refresh(token));
  const workspaceRecoveryPath=path.join(app.getPath('userData'),'workspace-recovery.json');
  const recoveryManager=createRecoveryManager({directory:app.getPath('userData'),storage,workspaceStorage});
  handle('getPendingRecoveries',()=>enqueueWrite(()=>recoveryManager.listPending()));
  handle('readPendingRecovery',id=>enqueueWrite(()=>recoveryManager.readPending(id)));
  handle('discardPendingRecovery',id=>enqueueWrite(()=>recoveryManager.discardPending(id)));
  handle('autosaveWorkspace',data=>enqueueWrite(()=>recoveryManager.writeWorkspace(data)));
  handle('getWorkspaceRecovery',()=>enqueueWrite(()=>workspaceStorage.read(workspaceRecoveryPath)));
  handle('discardWorkspaceRecovery',()=>enqueueWrite(()=>workspaceStorage.discard(workspaceRecoveryPath)));

  const preparedExports=new Map();
  handle('prepareExport',async request=>{
    if(!request||!['png','jpeg','webp','svg','pdf'].includes(request.format))throw new Error(UI_TEXT.errorExport);
    const bytes=request.format==='pdf'?await print.toPDF(request):storage.encodeExport(request.format,request.data);
    if(bytes.length>storage.MAX_BYTES)throw new Error(UI_TEXT.errorLarge);
    preparedExports.clear();const token=require('node:crypto').randomUUID();preparedExports.set(token,{bytes,format:request.format,name:storage.safeName(request.name)});
    return {token,bytes:bytes.length,format:request.format};
  });
  handle('savePreparedExport',token=>enqueueWrite(async()=>{
    const value=preparedExports.get(token);if(!value)throw new Error(UI_TEXT.errorExport);
    const extension=value.format==='jpeg'?'jpg':value.format;
    const selected=await dialog.showSaveDialog(window,{title:UI_TEXT.dialogExport,defaultPath:value.name+'.'+extension,filters:[{name:value.format.toUpperCase(),extensions:[extension]}]});
    if(selected.canceled||!selected.filePath)return null;
    if(path.extname(selected.filePath).toLowerCase()!=='.'+extension)throw new Error(UI_TEXT.errorExport);
    await protectSource(selected.filePath);await storage.atomicWrite(selected.filePath,value.bytes);
    preparedExports.delete(token);return {path:selected.filePath,bytes:value.bytes.length};
  }));
  handle('releasePreparedExport',token=>preparedExports.delete(token));

  handle('autosave', (data) => enqueueWrite(async () => {
    await recoveryManager.writeLegacy(data);
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
      action(UI_TEXT.menuSample, undefined, 'sample'),
      action(UI_TEXT.menuDocumentSettings, undefined, 'documentSettings'),
      action(UI_TEXT.menuSave, 'CmdOrCtrl+S', 'save'),
      action(UI_TEXT.menuSaveAs, 'CmdOrCtrl+Shift+S', 'saveAs'),
      action(UI_TEXT.menuSaveVersion, 'CmdOrCtrl+Alt+S', 'saveVersion'),
      { type: 'separator' },
      action(UI_TEXT.menuImport, 'CmdOrCtrl+I', 'import'),
      action(UI_TEXT.menuExport, 'CmdOrCtrl+Shift+E', 'export'),
      action(UI_TEXT.menuPrint, 'CmdOrCtrl+P', 'print'),
      { type: 'separator' },
      { label: UI_TEXT.menuClose, accelerator: 'Alt+F4', click: () => window?.close() }
    ] },
    { label: UI_TEXT.menuEdit, submenu: [
      action(UI_TEXT.menuUndo, 'CmdOrCtrl+Z', 'undo'),
      action(UI_TEXT.menuRedo, 'CmdOrCtrl+Shift+Z', 'redo'),
      action(UI_TEXT.menuFontLibrary, undefined, 'fontLibrary')
    ] },
    { label: UI_TEXT.menuView, submenu: [
      action(UI_TEXT.menuFit, 'CmdOrCtrl+0', 'fit'),
      action(UI_TEXT.menuActual, 'CmdOrCtrl+1', 'actual'),
      action(UI_TEXT.menuTogglePanels, undefined, 'togglePanels'),
      action(UI_TEXT.menuToggleRulers, 'CmdOrCtrl+R', 'toggleRulers'),
      action(UI_TEXT.menuToggleGuides, undefined, 'toggleGuides'),
      action(UI_TEXT.menuResetWorkspace, undefined, 'resetWorkspace'),
      { label: UI_TEXT.menuFullscreen, role: 'togglefullscreen' },
      ...(!app.isPackaged ? [{ label: UI_TEXT.menuDevTools, role: 'toggleDevTools' }] : [])
    ] },
    { label: UI_TEXT.menuHelp, submenu: [action(UI_TEXT.menuHelpGuide, 'F1', 'help')] }
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
  app.on('second-instance', (_event, argv, cwd) => { queueOpen(argv, cwd); if (window) { if (window.isMinimized()) window.restore(); window.focus(); } });
  app.whenReady().then(() => {
    session.defaultSession.setPermissionRequestHandler((contents, permission, callback) => callback(permission === 'local-fonts' && contents === window?.webContents && contents.getURL() === indexUrl));
    session.defaultSession.setPermissionCheckHandler((contents, permission) => permission === 'local-fonts' && contents === window?.webContents && contents.getURL() === indexUrl);
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
