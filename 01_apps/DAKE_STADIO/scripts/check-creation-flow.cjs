'use strict';
// Runs an unmodified app/EXE. Local main-process inspector replaces native dialog
// answers only; renderer controls and production IPC/storage/export remain real.
require('./local-test-env.cjs')();
const fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto'),net=require('node:net'),{spawn}=require('node:child_process'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),executable=path.resolve(process.argv[2]||require('electron'));
const packaged=!!process.argv[2]&&!executable.includes('node_modules');
const seed=path.resolve(process.argv[3]||path.join(root,'test-output','profile-seed-20260930'));
const output=path.join(root,'test-output','creation-flow-'+Date.now());fs.mkdirSync(output,{recursive:true});
const photo='C:/Windows/Web/Wallpaper/ThemeA/img20.jpg',family=process.env.STADIO_QA_FONT||'Zen Maru Gothic';
const hash=b=>crypto.createHash('sha256').update(b).digest('hex'),q=JSON.stringify,delay=ms=>new Promise(r=>setTimeout(r,ms));
function jpegDimensions(bytes){if(bytes.readUInt16BE(0)!==0xffd8)throw Error('Not JPEG');let offset=2;while(offset+4<bytes.length){if(bytes[offset++]!==255)throw Error('Invalid JPEG marker');while(bytes[offset]===255)offset++;const marker=bytes[offset++];if(marker===0xd9||marker===0xda)break;if(marker===0x01||marker>=0xd0&&marker<=0xd7)continue;const length=bytes.readUInt16BE(offset);if(length<2||offset+length>bytes.length)throw Error('Invalid JPEG segment');if([0xc0,0xc1,0xc2,0xc3,0xc5,0xc6,0xc7,0xc9,0xca,0xcb,0xcd,0xce,0xcf].includes(marker))return{width:bytes.readUInt16BE(offset+5),height:bytes.readUInt16BE(offset+3)};offset+=length;}throw Error('JPEG dimensions missing');}
const capacity=n=>n<1024?n+' B':n<1048576?(n/1024).toFixed(1)+' KiB':(n/1048576).toFixed(2)+' MiB';
const report={output,executable,packaged,executableSha256:hash(fs.readFileSync(executable)),profileSeed:seed,photoFixture:{path:photo,sha256:hash(fs.readFileSync(photo)),distribution:'Windows fixture used only in ignored test-output; do not bundle these test artworks.'},reports:[],method:[
 'EXE launched with local main-process inspector and renderer CDP; no production backdoor or source patch.',
 'Main inspector replaces Electron open/save dialogs with explicitly logged test paths. Native component, import, safe save and prepared-export IPC still execute.',
 'Renderer buttons/text and mask gestures use CDP mouse/key input. Document/geometry setup and source assertions use application APIs. Correction slider values dispatch input events.',
 'Copied mode clones every regular file from the supplied safe snapshot; user original settings and snapshot are never edited.',
 'Resume uses a new PID, STADIO_FONT_OFFLINE=1, blocked renderer network, and main HTTPS guard. No OS setting is changed.'
]};
if(packaged)report.appAsarSha256=hash(fs.readFileSync(path.join(path.dirname(executable),'resources','app.asar')));
function tree(dir,base=dir,result={}){for(const name of fs.readdirSync(dir)){const file=path.join(dir,name),stat=fs.lstatSync(file);if(stat.isSymbolicLink())throw Error('Seed has unsupported link: '+file);if(stat.isDirectory())tree(file,base,result);else if(stat.isFile())result[path.relative(base,file)]={bytes:stat.size,sha256:hash(fs.readFileSync(file))};}return result;}
const seedHashes=fs.existsSync(seed)?tree(seed):null;if(!seedHashes)throw Error('Safe profile snapshot is required: '+seed);
const snapshotManifest=seed+'-manifest.json';if(fs.existsSync(snapshotManifest))report.profileSnapshotManifest=snapshotManifest;
async function port(){const server=net.createServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));const value=server.address().port;await new Promise(r=>server.close(r));return value;}
async function endpoint(portValue,selector){const until=Date.now()+30000;while(Date.now()<until){try{const response=await fetch('http://127.0.0.1:'+portValue+'/json/list',{signal:AbortSignal.timeout(800)});const target=(await response.json()).find(selector);if(target)return target.webSocketDebuggerUrl;}catch{}await delay(80);}throw Error('Inspector unavailable '+portValue);}
async function socket(url,errors,consoleEvents=[]){const ws=new WebSocket(url),pending=new Map();let sequence=0;const rejectPending=()=>{for(const p of pending.values()){clearTimeout(p.timer);p.reject(Error('Protocol socket closed'));}pending.clear();};ws.addEventListener('close',rejectPending);ws.addEventListener('error',rejectPending);
 await new Promise((resolve,reject)=>{ws.addEventListener('open',resolve,{once:true});ws.addEventListener('error',reject,{once:true});});
 ws.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id&&pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);clearTimeout(p.timer);m.error?p.reject(Error(JSON.stringify(m.error))):p.resolve(m.result);}else if(m.method==='Runtime.exceptionThrown')errors.push(m.params.exceptionDetails);else if(m.method==='Runtime.consoleAPICalled'&&m.params.type==='error')consoleEvents.push(m.params.args.map(a=>a.description||a.value).join(' '));});
 return {command(method,params={}){return new Promise((resolve,reject)=>{if(ws.readyState!==1){reject(Error('Protocol socket is not open: '+method));return;}const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(Error('Protocol timeout '+method));},120000);pending.set(id,{resolve,reject,timer});ws.send(JSON.stringify({id,method,params}));});},
 close(){rejectPending();ws.close();}};
}
async function runMode(mode){
 const dir=path.join(output,mode),profile=path.join(dir,'profile');fs.mkdirSync(dir,{recursive:true});
 if(mode==='copied')fs.cpSync(seed,profile,{recursive:true});else fs.mkdirSync(profile);
 const r={mode,profile,checks:[],errors:[],console:[],processes:[],network:[],dialogResponses:[],screens:[],exports:{}};report.reports.push(r);
 const check=(name,value,detail)=>{r.checks.push({name,passed:!!value,...(detail===undefined?{}:{detail})});assert.ok(value,name);};
 if(mode==='copied')check('full safe profile copy matches all snapshot bytes',JSON.stringify(tree(profile))===JSON.stringify(seedHashes),{files:Object.keys(seedHashes).length,bytes:Object.values(seedHashes).reduce((sum,f)=>sum+f.bytes,0)});
 const logo=path.join(dir,'atelier-logo.dake'),card=path.join(dir,'business-card.dake');
 fs.writeFileSync(logo,JSON.stringify({format:'dake-stadio',version:2,name:'atelier DAKE logo',width:720,height:210,background:'transparent',fonts:[],meta:{dpi:300,unit:'px'},canvas:{version:'7.4.0',objects:[]}}));
 let child,main,renderer,lastPID,beforeRestartPNG;
 async function evalOn(channel,expression){const result=await channel.command('Runtime.evaluate',{expression,awaitPromise:true,returnByValue:true});if(result.exceptionDetails)throw Error(JSON.stringify(result.exceptionDetails));return result.result.value;}
 const progress=label=>{r.progress=label;fs.writeFileSync(path.join(dir,'progress.json'),JSON.stringify({progress:label,lastControl:r.lastControl,checks:r.checks,processes:r.processes},null,2));};
 const js=expression=>{r.lastExpression=expression.slice(0,300);progress('renderer '+r.lastExpression);return evalOn(renderer,expression);};
 const native=expression=>{progress('main '+expression.slice(0,180));return evalOn(main,expression);};
 const wait=async(expression,timeout=25000)=>{const until=Date.now()+timeout;while(Date.now()<until){if(await js(expression))return;await delay(50);}throw Error('Condition timeout '+expression);};
 async function click(selector){r.lastControl=selector;const point=await js('(()=>{const e=document.querySelector('+q(selector)+');if(!e)throw Error("Missing UI control");e.scrollIntoView({block:"nearest"});const r=e.getBoundingClientRect();if(!e.getClientRects().length||r.width<1||r.height<1)throw Error("UI control is hidden");const point={x:r.left+r.width/2,y:r.top+r.height/2};if(!e.contains(document.elementFromPoint(point.x,point.y)))throw Error("UI control is covered");window.__creationClickExpected=e;window.__creationClickHit=null;return point})()');await renderer.command('Input.dispatchMouseEvent',{type:'mouseMoved',...point,button:'none'});await renderer.command('Input.dispatchMouseEvent',{type:'mousePressed',...point,button:'left',buttons:1,clickCount:1});await renderer.command('Input.dispatchMouseEvent',{type:'mouseReleased',...point,button:'left',buttons:0,clickCount:1});if(!await js('window.__creationClickHit'))throw Error('Mouse click was not received by '+selector);}
 async function input(id,value){await click('#'+id);await renderer.command('Input.dispatchKeyEvent',{type:'keyDown',key:'a',code:'KeyA',modifiers:2,windowsVirtualKeyCode:65});await renderer.command('Input.dispatchKeyEvent',{type:'keyUp',key:'a',code:'KeyA',modifiers:2,windowsVirtualKeyCode:65});await renderer.command('Input.insertText',{text:String(value)});await renderer.command('Input.dispatchKeyEvent',{type:'keyDown',key:'Tab',code:'Tab'});await renderer.command('Input.dispatchKeyEvent',{type:'keyUp',key:'Tab',code:'Tab'});}
 async function range(id,value,event='input'){await js('(()=>{const e=document.getElementById('+q(id)+');e.value='+q(value)+';e.dispatchEvent(new Event('+q(event)+',{bubbles:true}));return true})()');}
 async function shot(name){const s=await renderer.command('Page.captureScreenshot',{format:'png'});const file=path.join(dir,name+'.png');fs.writeFileSync(file,Buffer.from(s.data,'base64'));r.screens.push(file);}
 async function launch(file,offline){
  const mainPort=await port(),renderPort=await port(),env={...process.env,STADIO_USER_DATA:profile,STADIO_FONT_OFFLINE:offline?'1':'0'};delete env.ELECTRON_RUN_AS_NODE;
  const args=[...(packaged?[]:[root]),'--inspect=127.0.0.1:'+mainPort,'--remote-debugging-port='+renderPort,file];
  const log=path.join(dir,offline?'offline-process.log':'online-process.log'),fd=fs.openSync(log,'w');try{child=spawn(executable,args,{cwd:dir,env,stdio:['ignore',fd,fd],windowsHide:true});}finally{fs.closeSync(fd);}r.processLogs=[...(r.processLogs||[]),log];child.once('exit',(code,signal)=>{r.processExits=[...(r.processExits||[]),{pid:child.pid,offline,code,signal}];});
  main=await socket(await endpoint(mainPort,()=>true),r.errors,r.console);
  await native('(()=>{const req=process.mainModule.require.bind(process.mainModule),electron=req("electron"),https=req("node:https");globalThis.__creationQA={open:null,save:null,dialogs:[],requests:[],offline:'+offline+'};const qa=globalThis.__creationQA;electron.dialog.showOpenDialog=async(_w,options)=>{qa.dialogs.push({kind:"open",title:options?.title,path:qa.open});return{canceled:!qa.open,filePaths:qa.open?[qa.open]:[]}};electron.dialog.showSaveDialog=async(_w,options)=>{qa.dialogs.push({kind:"save",title:options?.title,path:qa.save});return{canceled:!qa.save,filePath:qa.save||undefined}};const original=https.get;https.get=function(url,...args){qa.requests.push(String(url));if(qa.offline)throw Error("QA_OFFLINE_NETWORK");return original.call(this,url,...args)};return{pid:process.pid,offline:qa.offline};})()');
  for(let attempt=0;attempt<4;attempt++){try{progress('renderer attach '+renderPort);renderer=await socket(await endpoint(renderPort,t=>t.type==='page'&&t.url.includes('/build/index.html')),r.errors,r.console);await renderer.command('Runtime.enable');await wait('!!window.__studio?.workspace&&!!window.__studio?.imaging&&__studio.currentPath==='+q(file),40000);break;}catch(error){renderer?.close();if(!/socket|context.*destroyed/i.test(error.message)||attempt===3)throw error;r.reconnections=(r.reconnections||0)+1;}}
  if(offline){await renderer.command('Network.enable');await renderer.command('Network.emulateNetworkConditions',{offline:true,latency:0,downloadThroughput:0,uploadThroughput:0});}
  const pid=await native('process.pid');r.processes.push({pid,offline,file});if(lastPID)check('offline restart uses a new process ID',pid!==lastPID);lastPID=pid;
  await js('__studio.production.prepareNewText().then(()=>true)');await js('document.addEventListener("click",e=>{window.__creationClickHit=e.composedPath().includes(window.__creationClickExpected)},true);true');
  // Preserve snapshot content; dismiss recovery choices only inside this disposable clone.
  for(const id of ['dismiss-workspace','dismiss-recovery'])if(await js('!!document.getElementById('+q(id)+')')){await click('#'+id);await wait('!document.getElementById('+q(id)+')');}
 }
 async function setDialog(key,value){await native('__creationQA.'+key+'='+q(value));}
 async function close(){
  check('all production tabs are saved before normal close',await js('!__studio.workspace.anyDirty'));
  r.network.push(...await native('__creationQA.requests'));r.dialogResponses.push(...await native('__creationQA.dialogs'));
  const exited=new Promise(resolve=>child.once('exit',code=>resolve(code)));
  await native('(()=>{const w=process.mainModule.require("electron").BrowserWindow.getAllWindows().find(w=>w.webContents.getURL().includes("index.html"));setTimeout(()=>w.close(),0);return true})()');
  // Node inspector intentionally keeps Node alive while attached. Detach after the close request.
  main.close();renderer.close();const code=await Promise.race([exited,delay(15000).then(()=>{throw Error('Normal close timed out');})]);check('normal close exits successfully',code===0);main=renderer=null;
 }
 async function selectLayer(predicate){await js('(()=>{const e=__studio.engine,o=e.layers.find('+predicate+');if(!o)throw Error("Layer missing");e.select(o.id);return true})()');await wait('!!__studio.engine.selected');}
 async function glyphProof(label){
  const data=await js('(async()=>{const e=__studio.engine,o=e.selected,d=e.fonts.find(f=>f.family===o.fontFamily);if(!d?.data)throw Error("Missing embedded font");const alias="QA-flow-"+crypto.randomUUID(),face=new FontFace(alias,"url("+d.data+")",{weight:String(d.weightRange||d.weight||400),style:d.style||"normal"});await face.load();document.fonts.add(face);const render=async family=>{const c=await o.clone();c.set({fontFamily:family});c.initDimensions();return c.toCanvasElement().toDataURL()};const actual=await render(o.fontFamily),reference=await render(alias),fallback=await render("monospace");document.fonts.delete(face);return{actual,reference,fallback,family:o.fontFamily}})()');
  check(label+' glyph pixels equal independently loaded actual font bytes',data.actual===data.reference&&data.actual!==data.fallback,{family:data.family,actual:hash(data.actual),reference:hash(data.reference),fallback:hash(data.fallback)});
 }
 async function maskStroke(x,y,modeValue){
  await click('[data-panel="effects"]');await click('[data-im-target="layerMask"]');await wait('__studio.engine._imageEditingTarget==="layerMask"&&!__studio.engine.busy');
  await click(modeValue==='restore'?'#im-mask-restore':'#im-mask-hide');await range('im-mask-size',140);await range('im-mask-hardness',25);
  const coords=await js('(()=>{const e=__studio.engine,r=e.canvas.upperCanvasEl.getBoundingClientRect(),v=e.canvas.viewportTransform;return{x:r.left+v[4]+'+x+'*v[0],y:r.top+v[5]+'+y+'*v[3]}})()'),prior=await js('__studio.engine._historyIndex');
  await renderer.command('Input.dispatchMouseEvent',{type:'mousePressed',...coords,button:'left',buttons:1,clickCount:1});
  await renderer.command('Input.dispatchMouseEvent',{type:'mouseMoved',x:coords.x+35,y:coords.y+25,button:'left',buttons:1});
  await renderer.command('Input.dispatchMouseEvent',{type:'mouseReleased',x:coords.x+35,y:coords.y+25,button:'left',buttons:0,clickCount:1});await wait('!__studio.engine.busy&&!__studio.engine._strokeTask');
  check(modeValue+' brush commits one mask transaction',await js('__studio.engine._historyIndex=== '+(prior+1)));
 }
 try{
  await launch(logo,false);await click('[data-tool="text"]');await wait('__studio.engine.selected?.type==="textbox"');await input('prop-text','atelier DAKE');
  await js('__studio.engine.updateSelected({left:24,top:35,fontSize:50,fill:"#233f36"});true');
  await click('#choose-font');await click('#font-google');await wait('document.querySelector("#font-google").classList.contains("active")&&document.querySelectorAll(".font-row").length>0');await input('font-search',family);await wait('!!document.querySelector(".font-row")');
  await click('.font-preview-button');await wait('document.querySelector(".font-row").dataset.preview==="ready"',120000);await shot('01-font-preview');await click('.font-use-button');await wait('__studio.engine.selected.fontFamily==='+q(family));await click('#dialog-close');await glyphProof('logo');
  await js('(()=>{const e=__studio.engine,id=e.selected.id;e.addShape("rect",{left:24,top:122,width:140,height:8,rx:4,ry:4,fill:"#a78045",stroke:null,strokeWidth:0});e.select(id);return true})()');await click('#save');await wait('!__studio.dirty');const logoHash=hash(fs.readFileSync(logo));check('actual logo file stores editable text and font bytes',JSON.parse(fs.readFileSync(logo,'utf8')).fonts.some(f=>f.family===family&&f.data&&f.license));
  await shot('02-logo-saved');
  await js('(async()=>{await __studio.workspace.newDocument({width:1146,height:720,name:"atelier DAKE / business card",background:"#f4f0e6",meta:{dpi:300,unit:"mm",bleed:35,safe:47}});window.cardSession=__studio.workspace.active.id;return true})()');
  await setDialog('open',logo);await click('#place-component');await wait('!!document.getElementById("dialog-embedded")');await click('#dialog-embedded');await wait('__studio.engine.layers.some(o=>o.dakeComponent)');
  check('native component placement retains editable text and vectors',await js('__studio.engine.selected.dakeComponent.document.canvas.objects.some(o=>o.text==="atelier DAKE")&&__studio.engine.selected.getObjects().some(o=>o.type==="textbox")&&__studio.engine.selected.getObjects().some(o=>o.type==="rect")'));
  await js('__studio.engine.updateSelected({left:90,top:115,scaleX:.9,scaleY:.9});window.componentId=__studio.engine.selected.id;true');await click('[data-panel="edit"]');await click('#component-edit');await wait('__studio.workspace.editingChild');
  await selectLayer('o=>o.type==="textbox"');await wait('document.getElementById("component-parent-preview")?.naturalWidth>0&&document.querySelector(".busy-strip").hidden',40000);const oldPreview=await js('document.getElementById("component-parent-preview").src');await input('prop-text','atelier DAKE+');
  await wait('document.getElementById("component-parent-preview")?.naturalWidth>0&&document.getElementById("component-parent-preview").src!=='+q(oldPreview),40000);
  check('child editing gives live parent preview before applying',await js('__studio.workspace.sessions.find(s=>s.id===window.cardSession).data.canvas.objects.some(o=>o.dakeComponent?.document.canvas.objects.some(c=>c.text==="atelier DAKE"))'));
  await shot('03-child-live-preview');await click('#component-apply');await wait('!__studio.workspace.editingChild');check('child apply updates parent while original file remains unchanged',hash(fs.readFileSync(logo))===logoHash&&await js('__studio.engine.layers.some(o=>o.dakeComponent?.document.canvas.objects.some(c=>c.text==="atelier DAKE+"))'));
  await setDialog('open',photo);await click('#import');await wait('__studio.engine.selected?.type==="image"',40000);
  await js('(()=>{const e=__studio.engine,o=e.selected;window.photoId=o.id;window.photoSource=o.toObject().src;e.updateSelected({left:700,top:65,scaleX:500/o.width,scaleY:600/o.height});e.canvas.sendObjectToBack(o);e.commit("photo-layout");return true})()');
  await click('[data-panel="effects"]');await click('#image-adjust-open');await wait('!!document.getElementById("im-adjust-exposure")');await range('im-adjust-exposure',.25);await click('[data-im-page="color"]');await range('im-adjust-vibrance',.25);await click('[data-im-page="curve"]');
  const curve=await js('(()=>{const r=document.getElementById("im-tone-curve").getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height*.44}})()');
  await renderer.command('Input.dispatchMouseEvent',{type:'mousePressed',...curve,button:'left',buttons:1,clickCount:1});await renderer.command('Input.dispatchMouseEvent',{type:'mouseReleased',...curve,button:'left',buttons:0,clickCount:1});
  await wait('!__studio.engine._workerJob');await shot('04-photo-correction-preview');await click('#dialog-apply');await wait('!__studio.engine.imagingPreviewing');
  await maskStroke(740,300,'hide');await range('im-mask-feather',8);await range('im-mask-feather',8,'change');await wait('!__studio.engine.imagingPreviewing&&!__studio.engine._workerJob');
  check('photo correction and mask preserve imported original',await js('__studio.engine.selected.toObject().src===window.photoSource&&__studio.engine.selected.dakeImageEdit.exposure===.25&&__studio.engine.selected.dakeImageEdit.vibrance===.25&&__studio.engine.selected.dakeImageEdit.curve.some(([x,y])=>x!==y)&&__studio.engine.selected.dakeLayerMask.feather===8'));
  await click('[data-im-target="image"]');await click('[data-tool="text"]');await wait('__studio.engine.selected?.type==="textbox"');check('new card text inherits previously selected Google font',await js('__studio.engine.selected.fontFamily==='+q(family)));
  await click('[data-panel="edit"]');await input('prop-text','Designer / YUKI\nhello@example.invalid');check('card contact text is entered through the visible editor',await js('__studio.engine.selected.text==="Designer / YUKI\\nhello@example.invalid"'));await js('__studio.engine.updateSelected({left:110,top:390,fontSize:32,fill:"#233f36"});true');
  await setDialog('save',card);await click('#save');await wait('!__studio.dirty');beforeRestartPNG=await js('__studio.engine.exportRaster()');fs.writeFileSync(path.join(dir,'before-restart.png'),Buffer.from(beforeRestartPNG.split(',')[1],'base64'));
  await shot('05-card-saved');check('saved card contains editable component, image recipe and mask',JSON.parse(fs.readFileSync(card,'utf8')).canvas.objects.some(o=>o.dakeImageEdit&&o.dakeLayerMask));
  await close();await launch(card,true);
  const reopened=await js('__studio.engine.exportRaster()');fs.writeFileSync(path.join(dir,'after-restart.png'),Buffer.from(reopened.split(',')[1],'base64'));
  check('save close new-PID offline reopen renders identical PNG',reopened===beforeRestartPNG);beforeRestartPNG=null;
  await selectLayer('o=>o.dakeComponent');await click('[data-panel="edit"]');await click('#component-edit');await wait('__studio.workspace.editingChild');await selectLayer('o=>o.type==="textbox"');await wait('document.getElementById("component-parent-preview")?.naturalWidth>0&&document.querySelector(".busy-strip").hidden',40000);await glyphProof('offline reopened logo');
  const oldOfflinePreview=await js('document.getElementById("component-parent-preview").src');await input('prop-text','atelier DAKE / 2026');await wait('document.getElementById("component-parent-preview")?.naturalWidth>0&&document.getElementById("component-parent-preview").src!=='+q(oldOfflinePreview),40000);await click('#component-apply');await wait('!__studio.workspace.editingChild');await selectLayer('o=>o.type==="image"');await maskStroke(740,300,'restore');
  check('original logo and source photograph never changed',hash(fs.readFileSync(logo))===logoHash&&hash(fs.readFileSync(photo))===report.photoFixture.sha256);
  await click('#save');await wait('!__studio.dirty');await shot('06-reedited-offline');
  for(const format of ['png','jpeg','pdf']){
    const ext=format==='jpeg'?'jpg':format,file=path.join(dir,'business-card.'+ext);await setDialog('save',file);await click('#export');await wait('!!document.getElementById("export-format")');
    await range('export-format',format);await wait('!document.getElementById("dialog-export").disabled',60000);
    const info=await js('({size:document.getElementById("output-size").textContent,dimensions:document.getElementById("output-dimensions").textContent,note:document.getElementById("output-format-note").textContent})');
    check(format+' shows encoded capacity instead of estimate',/実測/.test(info.size));await shot('07-export-'+format);await click('#dialog-export');
    const until=Date.now()+20000;while(!fs.existsSync(file)&&Date.now()<until)await delay(50);assert.ok(fs.existsSync(file),'native prepared export wrote '+format);
    await wait('document.querySelector(".busy-strip").hidden');const bytes=fs.readFileSync(file);check(format+' measured capacity matches the final encoded file',info.size.endsWith(capacity(bytes.length)));r.exports[format]={file,bytes:bytes.length,sha256:hash(bytes),...info};
    if(format==='png')check('PNG output uses intended exact pixel dimensions',bytes.readUInt32BE(16)===1146&&bytes.readUInt32BE(20)===720);
    if(format==='jpeg'){const dimensions=jpegDimensions(bytes);check('JPEG output has real JPEG header and intended dimensions',bytes[0]===255&&bytes[1]===216&&bytes[2]===255&&dimensions.width===1146&&dimensions.height===720);r.exports[format].decodedDimensions=dimensions;}
    if(format==='pdf'){
      const text=bytes.toString('latin1'),m=/\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/.exec(text);
      check('PDF real page size equals artwork mm/dpi',!!m&&Math.abs(+m[1]-1146/300*72)<.000001&&Math.abs(+m[2]-720/300*72)<.000001);
      check('PDF structure and limitations are accurately shown',/\/Subtype \/Image/.test(text)&&/\/DeviceRGB/.test(text)&&!/\/Font\b/.test(text)&&/RGB/.test(info.note)&&/PDF\/X/.test(info.note));
    }
  }
  const offlineRequests=await native('__creationQA.requests');check('offline process performs no native font HTTP requests',offlineRequests.length===0);
  await close();check('no uncaught renderer exceptions',r.errors.length===0);check('no renderer error console messages',r.console.length===0);check('safe profile seed remains byte-identical',JSON.stringify(tree(seed))===JSON.stringify(seedHashes));r.passed=true;
 }catch(error){r.passed=false;r.failureProgress=r.progress;r.error=error.stack;try{await shot('failure');r.failureText=await js('document.body.innerText');r.failureState=await js('({busy:__studio.engine.busy,switching:__studio.workspace.switching,child:__studio.workspace.editingChild,processing:!document.querySelector(".busy-strip").hidden,previewWidth:document.getElementById("component-parent-preview")?.naturalWidth,receivedClick:window.__creationClickHit})');}catch{}try{r.dialogResponses.push(...await native('__creationQA.dialogs'));r.network.push(...await native('__creationQA.requests'));}catch{}}
 finally{main?.close();renderer?.close();if(child&&child.exitCode===null)child.kill();fs.writeFileSync(path.join(dir,'report.json'),JSON.stringify(r,null,2));}
}
(async()=>{await runMode('fresh');await runMode('copied');report.passed=report.reports.every(r=>r.passed);report.finished=new Date().toISOString();const file=path.join(output,'report.json');fs.writeFileSync(file,JSON.stringify(report,null,2));const evidence=path.join(root,'evidence','creation-flow-'+(packaged?'packaged':'source')+'-results.json');fs.writeFileSync(evidence,JSON.stringify(report,null,2));console.log(JSON.stringify({passed:report.passed,output,evidence,runs:report.reports.map(r=>({mode:r.mode,passed:r.passed,checks:r.checks.length,error:r.error,step:r.failureProgress}))},null,2));process.exitCode=report.passed?0:1;})().catch(error=>{console.error(error);process.exitCode=1;});

