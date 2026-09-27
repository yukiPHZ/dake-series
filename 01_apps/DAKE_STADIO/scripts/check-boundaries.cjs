
'use strict';
require('./local-test-env.cjs')();
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const {app,BrowserWindow,dialog,ipcMain}=require('electron');
const root=path.resolve(__dirname,'..'),output=path.join(root,'test-output','boundaries-'+Date.now());
fs.mkdirSync(output,{recursive:true});process.env.STADIO_USER_DATA=path.join(output,'profile');
const report={output,started:new Date().toISOString(),checks:[],errors:[]};
let saveTarget,openTargets=[],win;
dialog.showSaveDialog=async()=>({canceled:!saveTarget,filePath:saveTarget});
dialog.showOpenDialog=async()=>({canceled:!openTargets.length,filePaths:openTargets});
dialog.showErrorBox=(title,message)=>{throw new Error(title+': '+message);};
// Only gate IPC responses after the real production handler finishes its file work.
const gates=new Map(),nativeHandle=ipcMain.handle.bind(ipcMain);
ipcMain.handle=(channel,callback)=>nativeHandle(channel,async(...args)=>{
 const gate=gates.get(channel),result=await callback(...args);
 if(gate){gate.entered=true;await gate.promise;}return result;
});
const gate=channel=>{let release;const value={entered:false,promise:new Promise(resolve=>release=resolve),release:()=>{gates.delete(channel);release();}};gates.set(channel,value);return value;};
require('node:https').get=()=>{throw new Error('BOUNDARY_CHECK_OFFLINE');};
require('../desktop/main.cjs');
const delay=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const js=source=>win.webContents.executeJavaScript(source,true);
const json=value=>JSON.stringify(value);
function check(name,value,detail){assert.ok(value,name);report.checks.push({name,passed:true,...(detail?{detail}:{})});}
async function until(predicate,timeout=15000){const end=Date.now()+timeout;while(Date.now()<end){if(await predicate())return;await delay(25);}throw new Error('Boundary wait timed out');}
async function wait(source,timeout){return until(()=>js(source),timeout);}
async function click(id){await js('document.getElementById('+json(id)+').click();true');await delay(30);}
async function startNew(name,preset='logo'){
 await js('window.boundaryNewDone=false;window.boundaryNewError=null;__studio.production.newDocument('+json(preset)+').then(()=>window.boundaryNewDone=true).catch(error=>{window.boundaryNewError=error.message;window.boundaryNewDone=true});true');
 await wait("!!document.getElementById('dialog-discard')||!!document.getElementById('new-width')");
 if(await js("!!document.getElementById('dialog-discard')"))await click('dialog-discard');
 await wait("!!document.getElementById('new-width')");
 await js("document.getElementById('new-name').value="+json(name)+";true");await click('dialog-create');
 await wait('window.boundaryNewDone');check('new document command completes '+name,await js('!window.boundaryNewError'));
}
async function startSave(as=true){
 await js('window.boundarySaveDone=false;window.boundarySaveResult=null;window.boundarySaveError=null;__studio.save('+as+').then(value=>{window.boundarySaveResult=value;window.boundarySaveDone=true}).catch(error=>{window.boundarySaveError=error.message;window.boundarySaveDone=true});true');
}
async function finishSave(){await wait('window.boundarySaveDone');check('native save completes without exception',await js('!window.boundarySaveError'));}
function fontFixtures(){
 const base=path.join(root,'test-output');
 const directories=fs.readdirSync(base).filter(name=>/^production-/.test(name)).sort().reverse();
 for(const directory of directories){
  const fonts=path.join(base,directory,'profile-prepare','fonts'),folder=path.join(fonts,'descriptors');if(!fs.existsSync(folder))continue;
  const descriptors=fs.readdirSync(folder).filter(name=>name.endsWith('.json')).map(name=>JSON.parse(fs.readFileSync(path.join(folder,name),'utf8')));
  const noto=descriptors.find(d=>d.family==='Noto Sans JP'),zen=descriptors.find(d=>d.family==='Zen Kaku Gothic New');
  if(!noto||!zen)continue;
  return[noto,zen].map((d,i)=>({...d,id:'boundary-version-'+i,family:'Boundary Font',weight:700,style:'normal',
   license:fs.readFileSync(path.join(fonts,'licenses',d.licenseSha256+'.txt'),'utf8'),
   data:'data:font/ttf;base64,'+fs.readFileSync(path.join(fonts,'blobs',d.sha256+'.sfnt')).toString('base64')}));
 }
 throw new Error('Two existing OFL fixture fonts were not found');
}
async function flow(){
 await app.whenReady();win=BrowserWindow.getAllWindows()[0];
 win.webContents.on('console-message',event=>{if(event.level==='error')report.errors.push(event.message);});
 if(win.webContents.isLoading())await new Promise(resolve=>win.webContents.once('did-finish-load',resolve));await wait('!!window.__studio');
 await startNew('Boundary card','card');
 const created=await js('({dpi:__studio.engine.meta.dpi,unit:__studio.engine.meta.unit,width:__studio.engine.width,history:__studio.engine._history.length,undo:__studio.engine.canUndo})');
 check('new card metadata and dimensions are the initial single history state',created.dpi===300&&created.unit==='mm'&&created.history===1&&!created.undo,created);
 await js('__studio.engine.undo();true');
 check('undo immediately after creation cannot change PDF physical size',await js("!__studio.engine.canUndo&&__studio.engine.meta.dpi===300&&Math.abs(__studio.engine.width/300*25.4-97)<.1"));
 await js("__studio.engine.addShape('rect',{fill:'#f00'});true");
 saveTarget=path.join(output,'old-card.dake');const oldPath=saveTarget,saveGate=gate('stadio:saveProject');
 await startSave();await until(()=>saveGate.entered);
 check('delayed save really wrote original project before response gate',JSON.parse(fs.readFileSync(oldPath,'utf8')).name==='Boundary card');
 await startNew('Work after old save');
 saveGate.release();await finishSave();
 check('old save response cannot attach its path or mark the new work saved',await js("__studio.engine.name==='Work after old save'&&__studio.currentPath===null&&__studio.dirty&&window.boundarySaveResult===false"));
 saveTarget=path.join(output,'current-work.dake');await startSave();await finishSave();
 check('new work subsequently saves to its own path',await js('__studio.currentPath==='+json(saveTarget)+'&&!__studio.dirty'));
 await js("__studio.engine.addShape('ellipse');true");
 const editGate=gate('stadio:saveProject');await startSave(false);await until(()=>editGate.entered);
 await js("__studio.engine.addShape('triangle');true");editGate.release();await finishSave();
 check('edits during save remain dirty and save returns false',await js("__studio.engine.layers.length===2&&__studio.dirty&&window.boundarySaveResult===false"));
 const ensureGate=gate('stadio:saveProject');
 await js("window.boundaryEnsureDone=false;__studio.production.newDocument('logo').then(()=>window.boundaryEnsureDone=true);true");
 await wait("!!document.getElementById('dialog-save')");await click('dialog-save');await until(()=>ensureGate.entered);
 await js("__studio.engine.addShape('rect');true");ensureGate.release();await wait('window.boundaryEnsureDone');
 check('ensureSaved cancels work replacement when further edits arrive during saving',await js("__studio.engine.name==='Work after old save'&&__studio.dirty&&__studio.engine.layers.length===3&&!document.getElementById('new-width')"));
 const recoveryGate=gate('stadio:discardRecovery');await startSave(false);await until(()=>recoveryGate.entered);
 await startNew('Work after recovery discard');
 recoveryGate.release();await finishSave();
 check('late recovery discard response cannot clear new document dirty state',await js("__studio.engine.name==='Work after recovery discard'&&__studio.currentPath===null&&__studio.dirty&&window.boundarySaveResult===false"));
 saveTarget=path.join(output,'before-open.dake');await startSave();await finishSave();
 openTargets=[oldPath];const openGate=gate('stadio:openProject');
 await js("window.boundaryOpenDone=false;window.boundaryOpenError=null;__studio.open().then(()=>window.boundaryOpenDone=true).catch(error=>{window.boundaryOpenError=error.message;window.boundaryOpenDone=true});true");
 await until(()=>openGate.entered);await js("__studio.engine.addShape('triangle');window.beforeOpenResponse=JSON.stringify(__studio.engine.serialize());true");
 openGate.release();await wait('window.boundaryOpenDone');
 check('edits during native file reading abort replacement safely',await js("window.boundaryOpenError==='DOCUMENT_CHANGED'&&JSON.stringify(__studio.engine.serialize())===window.beforeOpenResponse&&__studio.dirty"));

 const fixtures=fontFixtures();report.fontFixtures=fixtures.map(({family,sha256,weight,style})=>({family,sha256,weight,style}));
 check('font versions share family weight style but have different real binary hashes',fixtures[0].sha256!==fixtures[1].sha256&&fixtures.every(f=>f.family==='Boundary Font'&&f.weight===700&&f.style==='normal'));
 await js('window.boundaryFonts='+json(fixtures)+';true');
 await js("__studio.engine.newDocument({width:700,height:250,background:'transparent',name:'font versions'});__studio.engine.addText('日本語と AVATAR');__studio.engine.updateSelected({left:25,top:25,fontSize:42,fill:'#000000'});true");
 await js('(async()=>{await __studio.production.useFont(window.boundaryFonts[0]);window.oldFontFace=[...document.fonts].find(face=>face.family.replace(/[\"\\\']/g,\"\")===\"Boundary Font\");window.oldFontPNG=await __studio.engine.exportRaster();window.oldFontProject=__studio.engine.serialize();return true;})()');
 await js('(async()=>{await __studio.production.useFont(window.boundaryFonts[1]);window.newFontFace=[...document.fonts].find(face=>face.family.replace(/[\"\\\']/g,\"\")===\"Boundary Font\");window.newFontPNG=await __studio.engine.exportRaster();return true;})()');
 check('different embedded font version changes actual rendering',await js("window.oldFontPNG!==window.newFontPNG&&window.oldFontFace!==window.newFontFace&&document.fonts.has(window.newFontFace)&&!document.fonts.has(window.oldFontFace)"));
 await js('__studio.engine.undo();true');await wait('!__studio.engine.busy');
 check('undo restores earlier same-variant hash and actual FontFace',await js("document.fonts.has(window.oldFontFace)&&!document.fonts.has(window.newFontFace)&&__studio.engine.fonts[0].sha256===window.boundaryFonts[0].sha256"));
 check('undo restores exact old font pixels',await js('(async()=>await __studio.engine.exportRaster()===window.oldFontPNG)()'));
 await js('__studio.engine.redo();true');await wait('!__studio.engine.busy');
 check('redo restores newer same-variant Face and pixels',await js('(async()=>document.fonts.has(window.newFontFace)&&!document.fonts.has(window.oldFontFace)&&await __studio.engine.exportRaster()===window.newFontPNG)()'));
 await js("window.beforeBroken=JSON.stringify(__studio.engine.serialize());window.beforeBrokenPath=__studio.currentPath;window.beforeBrokenDirty=__studio.dirty;window.brokenProject=structuredClone(window.oldFontProject);window.brokenProject.canvas.objects.push({type:'Image',src:'data:image/png;base64,broken',width:20,height:20});true");
 await js("(async()=>{window.brokenRejected=false;try{await __studio.production.loadDocument(window.brokenProject)}catch{window.brokenRejected=true;}return true;})()");
 check('corrupt image after real old-font restoration preserves current project and state',await js("window.brokenRejected&&JSON.stringify(__studio.engine.serialize())===window.beforeBroken&&__studio.currentPath===window.beforeBrokenPath&&__studio.dirty===window.beforeBrokenDirty&&!__studio.engine.busy"));
 check('corrupt load restores prior active FontFace and exact current pixels',await js("(async()=>document.fonts.has(window.newFontFace)&&!document.fonts.has(window.oldFontFace)&&await __studio.engine.exportRaster()===window.newFontPNG)()"));

 await js("window.preparedProject=structuredClone(window.oldFontProject);window.preparedProject.name='Prepared document';window.preparedProject.fonts[0].family='Boundary Pending Font';window.preparedProject.canvas.objects[0].fontFamily='Boundary Pending Font';window.nativeFaceLoad=FontFace.prototype.load;FontFace.prototype.load=function(){const face=this;return new Promise((resolve,reject)=>{window.releaseBoundaryFont=()=>window.nativeFaceLoad.call(face).then(resolve,reject)})};window.prepareDone=false;window.prepareError=null;__studio.production.loadDocument(window.preparedProject).then(()=>window.prepareDone=true).catch(error=>{window.prepareError=error.message;window.prepareDone=true});true");
 await wait('!!window.releaseBoundaryFont');check('font preparation runs inside engine busy guard',await js('__studio.engine.busy'));
 const guarded=await js("(async()=>{let add=false,undo=false,newWork=false;try{__studio.engine.addShape('rect')}catch(error){add=error.message==='BUSY'}try{await __studio.engine.undo()}catch(error){undo=error.message==='BUSY'}try{__studio.engine.newDocument({width:10,height:10})}catch(error){newWork=error.message==='BUSY'}return{add,undo,newWork,unchanged:JSON.stringify(__studio.engine.serialize())===window.beforeBroken};})()");
 check('editing undo and new work are rejected throughout font preparation',guarded.add&&guarded.undo&&guarded.newWork&&guarded.unchanged,guarded);
 await js('FontFace.prototype.load=window.nativeFaceLoad;window.releaseBoundaryFont();true');await wait('window.prepareDone');
 check('released font preparation completes a single safe document swap',await js("!window.prepareError&&!__studio.engine.busy&&__studio.engine.name==='Prepared document'&&__studio.engine._history.length===1"));
 const closeGate=gate('stadio:discardRecovery');
 win.close();await wait("!!document.getElementById('dialog-discard')");await click('dialog-discard');await until(()=>closeGate.entered);
 await js("__studio.engine.addShape('rect',{fill:'#123456'});true");closeGate.release();await delay(100);
 check('actual native close is cancelled when edits arrive during recovery discard',!win.isDestroyed());
 check('cancelled close keeps latest work open and dirty',await js("__studio.engine.name==='Prepared document'&&__studio.engine.layers.length===2&&__studio.dirty"));
 report.passed=true;report.finished=new Date().toISOString();
 fs.writeFileSync(path.join(output,'result.json'),JSON.stringify(report,null,2));fs.writeFileSync(path.join(root,'evidence/boundary-check-results.json'),JSON.stringify(report,null,2));
 console.log(JSON.stringify({passed:true,checks:report.checks.length,output}));app.exit(0);
}
app.on('before-quit',event=>{if(!report.passed){event.preventDefault();report.passed=false;report.failure='Application attempted to quit before boundary checks completed';fs.writeFileSync(path.join(output,'result.json'),JSON.stringify(report,null,2));fs.writeFileSync(path.join(root,'evidence/boundary-check-results.json'),JSON.stringify(report,null,2));app.exit(1);}});
flow().catch(error=>{report.passed=false;report.failure=error.stack;fs.writeFileSync(path.join(output,'result.json'),JSON.stringify(report,null,2));fs.writeFileSync(path.join(root,'evidence/boundary-check-results.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1);});

