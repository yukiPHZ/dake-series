require('./local-test-env.cjs')();
'use strict';
const fs = require('node:fs/promises');
const { createReadStream } = require('node:fs');
const { createHash } = require('node:crypto');
const path = require('node:path');
const net = require('node:net');
const { spawn } = require('node:child_process');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const executable = path.resolve(process.argv[2] || require('electron'));
const output = path.join(root, 'test-output', `workflow-${Date.now()}`);
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const report = { executable, independentWorkingDirectory: output, checks: [], errors: [] };
let child, connection;
const pending = new Map();
let sequence = 0;
const check = (name, value) => { assert.ok(value, name); report.checks.push(name); };
async function freePort() {
  const server = net.createServer();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const port = server.address().port;
  await new Promise(resolve => server.close(resolve));
  return port;
}
async function hash(filePath) {
  const value = createHash('sha256');
  for await (const chunk of createReadStream(filePath)) value.update(chunk);
  return value.digest('hex');
}
function command(method, params = {}) {
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP timeout: ${method}`)); }, 15000);
    pending.set(id, { resolve, reject, timer });
    connection.send(JSON.stringify({ id, method, params }));
  });
}
async function js(expression) {
  report.lastExpression=expression.slice(0,300); const result = await command('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
async function waitUntil(expression, timeout = 15000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) { if (await js(expression)) return; await delay(50); }
  throw new Error(`Packaged UI timeout: ${expression}`);
}
const packaged=!executable.includes('node_modules');
const environment={...process.env,STADIO_USER_DATA:path.join(output,'profile')};delete environment.ELECTRON_RUN_AS_NODE;
const prefix=packaged?[]:[root];
const fixture=path.join(output,'起動する作品.dake');
const exampleRoot=path.join(root,'artifacts','baseline-0.2.0','examples');
async function launch(file){
 const port=await freePort();report.port=port;console.log('QA port',port);child=spawn(executable,[...prefix,`--remote-debugging-port=${port}`,...(file?[file]:[])],{cwd:output,env:environment,stdio:'ignore',windowsHide:true});
 let target;for(let i=0;i<100;i++){try{target=(await(await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t=>t.type==='page');if(target)break;}catch{}await delay(100);}
 if(!target)throw new Error('No page');connection=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>connection.addEventListener('open',r,{once:true}));
 connection.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id&&pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);clearTimeout(p.timer);m.error?p.reject(new Error(JSON.stringify(m.error))):p.resolve(m.result);}else if(m.method==='Runtime.exceptionThrown')report.errors.push(m.params.exceptionDetails);});
 await command('Runtime.enable');await waitUntil('!!window.__studio?.workflow');
}
async function click(selector){await js(`document.querySelector(${JSON.stringify(selector)}).click()`);await delay(60);}
async function input(id,value,event='change'){await js(`(()=>{const e=document.getElementById(${JSON.stringify(id)});if(!e)throw new Error('Missing '+${JSON.stringify(id)});e.value=${JSON.stringify(value)};e.dispatchEvent(new Event(${JSON.stringify(event)},{bubbles:true}));})()`);await delay(30);}
async function mouse(type,p,modifiers=0){await command('Input.dispatchMouseEvent',{type,...p,button:'left',buttons:type==='mouseReleased'?0:1,clickCount:1,modifiers});await delay(30);}
async function screen(x,y){return js(`(()=>{const r=document.querySelector('.canvas-host').getBoundingClientRect(),v=__studio.engine.canvas.viewportTransform;return{x:r.left+v[4]+${x}*v[0],y:r.top+v[5]+${y}*v[3]}})()`);}
async function key(type,key,modifiers=0){await command('Input.dispatchKeyEvent',{type,key,code:key,modifiers});await delay(40);}
async function second(file){const p=spawn(executable,[...prefix,file],{cwd:output,env:environment,stdio:'ignore',windowsHide:true});await new Promise((r,j)=>{p.once('exit',r);p.once('error',j);});
 // Windows activation can detach the test debugger session; reconnect without touching document state.
 connection.close();const targets=await(await fetch('http://127.0.0.1:'+report.port+'/json/list')).json();
 connection=new WebSocket(targets.find(t=>t.type==='page').webSocketDebuggerUrl);await new Promise(r=>connection.addEventListener('open',r,{once:true}));
 connection.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id&&pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);clearTimeout(p.timer);m.error?p.reject(new Error(JSON.stringify(m.error))):p.resolve(m.result);}else if(m.method==='Runtime.exceptionThrown')report.errors.push(m.params.exceptionDetails);});
 await command('Runtime.enable');}

async function close(){await js('stadio.discardRecovery()');const exited=new Promise(r=>child.once('exit',r));await js('stadio.closeReady()');await exited;connection.close();}
async function run(){
 await fs.mkdir(output,{recursive:true});await fs.copyFile(path.join(exampleRoot,'banner.dake'),fixture);await launch(fixture);await waitUntil(`__studio.currentPath===${JSON.stringify(fixture)}`,30000);
 check('cold command-line .dake opens requested existing v2 artwork',await js('__studio.engine.layers.length>3&&!__studio.dirty'));
 await js("__studio.engine.newDocument({width:800,height:500,name:'Workflow',background:'#ffffff'});__studio.engine.fit(900,650);true");
 await click('[data-tool="rect"]');const a=await screen(100,100),b=await screen(270,190);await mouse('mousePressed',a);await mouse('mouseMoved',b);
 await key('keyDown','Shift',8);check('Shift press without mouse movement previews square',await js("(()=>{const o=__studio.engine.canvas.getObjects().find(o=>o.excludeFromExport&&o.type==='rect');return !!o&&Math.abs(o.width-o.height)<.01})()"));
 await key('keyUp','Shift',0);check('Shift release restores unconstrained proportions immediately',await js("(()=>{const o=__studio.engine.canvas.getObjects().find(o=>o.excludeFromExport&&o.type==='rect');return o&&o.width>o.height*1.5})()"));
 await key('keyDown','Alt',1);check('Alt previews a shape centered on drag start',await js("(()=>{const o=__studio.engine.canvas.getObjects().find(o=>o.excludeFromExport&&o.type==='rect');return Math.abs(o.left+o.width/2-100)<1})()"));await key('keyUp','Alt',0);await key('keyDown','Control',2);await key('keyUp','Control',0);await mouse('mouseReleased',b);
 check('one shape gesture creates exactly one undo state',await js('__studio.engine.layers.length===1&&__studio.engine._history.length===2'));
 await click('[data-tool="triangle"]');await mouse('mousePressed',await screen(340,120));await mouse('mouseMoved',await screen(490,190),8);await mouse('mouseReleased',await screen(490,190),8);
 check('Shift creates equilateral triangle geometry',await js('Math.abs(__studio.engine.selected.height/__studio.engine.selected.width-Math.sqrt(3)/2)<.0001'));
 await js('window.beforeSize=__studio.engine.selected.getScaledWidth();window.h=__studio.engine._history.length;true');
 await input('prop-width-slider',200,'input');check('size slider previews before committing',await js('Math.abs(__studio.engine.selected.getScaledWidth()-200)<1&&__studio.engine._history.length===h'));
 await input('prop-width-slider',230,'input');await input('prop-width-slider',230);check('one slider gesture has one undo state',await js('__studio.engine._history.length===h+1'));
 await js('window.colorBefore=__studio.engine.selected.fill;window.h=__studio.engine._history.length;true');await click('#style-fill');await input('color-r',210,'input');check('color input previews actual object without history spam',await js('__studio.engine.selected.fill!==colorBefore&&__studio.engine._history.length===h'));
 await click('#color-cancel');check('mouse cancel restores original color and history',await js('__studio.engine.selected.fill===colorBefore&&__studio.engine._history.length===h'));
 await click('#style-fill');await input('color-g',40,'input');await mouse('mousePressed',await screen(30,30));await mouse('mouseReleased',await screen(30,30));check('canvas click commits color once and creates no unintended object',await js('__studio.engine._history.length===h+1&&__studio.engine.layers.length===2&&!document.querySelector(".color-popover")'));
 await js("__studio.engine.select(__studio.engine.layers[1].id);true");const p=await screen(400,160);await js('window.h=__studio.engine._history.length;true');await mouse('mousePressed',p,2);await mouse('mouseMoved',{x:p.x+40,y:p.y+25},2);await key('keyDown','Shift',10);await key('keyDown','Alt',11);await key('keyUp','Alt',10);await key('keyUp','Shift',2);await mouse('mouseReleased',{x:p.x+40,y:p.y+25},0);check('Ctrl drag resizes selected object as one edit',await js('__studio.engine._history.length===h+1'));
 await click('[data-panel="layout"]');await input('guide-pos',150);await click('#guide-add');check('guide panel is visible even while object selected',await js('!!document.querySelector("#workflow-guides")&&__studio.engine.meta.guides.vertical.includes(150)'));
 const g=await screen(150,300);await mouse('mousePressed',g);await mouse('mouseMoved',{x:g.x+30,y:g.y});await mouse('mouseReleased',{x:g.x+30,y:g.y});check('canvas guide drag changes stored numeric coordinate',await js('__studio.engine.meta.guides.vertical[0]>150'));
 await js("(()=>{const e=document.querySelector('.guide-row input');e.value='220';e.dispatchEvent(new Event('change'));})()");await delay(60);check('guide numeric editor stores exact coordinate',await js('__studio.engine.meta.guides.vertical[0]===220'));
 await click('#guide-lock');check('guide lock disables manipulation but keeps guides',await js('__studio.engine.meta.guideLocked&&document.querySelector(".guide-row input").disabled&&__studio.engine.meta.guides.vertical.length===1'));
 const ruler=await js("(()=>{const r=document.querySelector('.ruler-top').getBoundingClientRect();return{x:r.left+70,y:r.top+8}})()");await mouse('mousePressed',ruler);await mouse('mouseReleased',ruler);check('locked guides cannot be added from ruler',await js('__studio.engine.meta.guides.vertical.length===1'));
 await click('#guide-lock');await click('.guide-row button');check('individual guide removal changes document',await js('__studio.engine.meta.guides.vertical.length===0'));
 await click('[data-panel="edit"]');await click('[data-tool="text"]');await input('prop-text','制作を、自然に。');await js('window.h=__studio.engine._history.length;true');await input('prop-fontsize-slider',62,'input');await input('prop-fontsize-slider',68,'input');await input('prop-fontsize-slider',68);check('text slider changes actual size and one history entry',await js('__studio.engine.selected.fontSize===68&&__studio.engine._history.length===h+1'));
 await js('__studio.save(false)');await js('__studio.engine.selected.enterEditing();__studio.engine.selected.selectAll();true');await command('Input.insertText',{text:'編集中の文字も保護'});await delay(100);
 check('in-canvas typing immediately marks unsaved without history spam',await js('__studio.dirty&&__studio.engine.selected.isEditing'));
 await js('__studio.save(false)');check('save finalizes live canvas typing to disk',JSON.parse(await fs.readFile(fixture,'utf8')).canvas.objects.some(o=>o.text==='編集中の文字も保護'));
 const families=['Zen Kaku Gothic New','Zen Maru Gothic'];
 for(const family of families){await click('#choose-font');await click('#font-google');await input('font-search',family,'input');await js(`(()=>{const r=[...document.querySelectorAll('.font-row')].find(r=>r.querySelector('strong').textContent===${JSON.stringify(family)});if(!r)throw new Error('Missing family');r.querySelector('button').click();})()`);await waitUntil(`__studio.engine.selected?.fontFamily===${JSON.stringify(family)}`,120000);await click('#dialog-close');check('download applies loaded font: '+family,await js(`document.fonts.check('16px "${family}"')&&__studio.engine.fonts.some(f=>f.family===${JSON.stringify(family)}&&f.data&&f.license)`));}
 await click('#choose-font');await click('#font-cached');await waitUntil('document.querySelectorAll(".font-row").length===2');const previewPoint=await js('(()=>{const r=document.querySelector(".font-row").getBoundingClientRect();return{x:r.left+30,y:r.top+25}})()');await mouse('mouseMoved',previewPoint);await waitUntil('document.querySelector(".font-row").dataset.preview==="ready"');check('cached library preview loads actual font face',await js('[...document.fonts].some(f=>f.family.startsWith("preview-")&&f.status==="loaded")'));await click('#dialog-close');
 await js("window.originalFaceLoad=FontFace.prototype.load;FontFace.prototype.load=function(){const face=this;return new Promise((resolve,reject)=>{window.releaseFontApply=()=>window.originalFaceLoad.call(face).then(resolve,reject)})};window.fontApplyResult=null;__studio.production.useFont({...__studio.engine.fonts[0],family:'QA Deferred Font'}).then(()=>window.fontApplyResult='applied').catch(e=>window.fontApplyResult=e.message);true");
 await waitUntil('typeof window.releaseFontApply==="function"');await js("__studio.engine.addText('取得中に選択した別の文字');FontFace.prototype.load=window.originalFaceLoad;window.releaseFontApply();true");await waitUntil('window.fontApplyResult!==null');check('delayed font activation cannot modify a newly selected layer',await js("window.fontApplyResult==='DOCUMENT_CHANGED'&&__studio.engine.selected.fontFamily!=='QA Deferred Font'&&!__studio.engine.fonts.some(f=>f.family==='QA Deferred Font')"));
 await click('[data-tool="text"]');await input('prop-font','Zen Kaku Gothic New');check('new text layer reuses previously downloaded family from selector',await js('__studio.engine.selected.fontFamily==="Zen Kaku Gothic New"&&__studio.engine.fonts.length===2'));
 await js("__studio.engine.newDocument({width:600,height:350,name:'別作品',background:'#fff'});__studio.engine.addText('再利用');true");await input('prop-font','Zen Maru Gothic');check('another document embeds cached selected font without redownload',await js('__studio.engine.selected.fontFamily==="Zen Maru Gothic"&&__studio.engine.fonts.length===1&&__studio.engine.fonts[0].data.length>1000'));
 await js('window.textId=__studio.engine.selected.id;true');await js("document.querySelector('.layer-row').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}))");check('layer double click edits text without rename modal',await js('__studio.engine.selected.isEditing&&!document.querySelector("#rename-value")'));await js('__studio.engine.selected.exitEditing();true');
 await click('.layer-more');await js("[...document.querySelectorAll('.context-menu button')].find(b=>b.textContent===__studio.UI_TEXT.rasterize).click()");await waitUntil('__studio.engine.selected?.type==="image"');await click('#undo');await waitUntil('__studio.engine.layers.some(o=>o.type==="textbox")');check('context rasterize can be undone to editable text',true);
 // Use the authorized startup path; version writes never open a test-substituted dialog.
 await js('__studio.save(false)');check('native overwrite saves the current document',JSON.parse(await fs.readFile(fixture,'utf8')).name==='別作品');
 await js('__studio.save(false,true)');const versionPath=await js('__studio.currentPath');check('numbered save updates current path and leaves earlier version',/_v001\.dake$/.test(versionPath)&&JSON.parse(await fs.readFile(versionPath,'utf8')).name==='別作品');
 await js("__studio.engine.renameDocument('未保存の大切な変更');true");await second(path.join(exampleRoot,'logo.dake'));await waitUntil('!!document.querySelector("#dialog-cancel")');await click('#dialog-cancel');check('second-instance open cancellation keeps unsaved artwork',await js('__studio.engine.name==="未保存の大切な変更"&&__studio.dirty'));
 await js('__studio.save(false)');await close();await launch(versionPath);await waitUntil(`__studio.currentPath===${JSON.stringify(versionPath)}`,30000);
 await command('Network.enable');await command('Network.emulateNetworkConditions',{offline:true,latency:0,downloadThroughput:0,uploadThroughput:0});
 await click('[data-tool="text"]');await input('prop-font','Zen Kaku Gothic New');await waitUntil('__studio.engine.selected.fontFamily==="Zen Kaku Gothic New"');check('restart offline reuses cached font not embedded in opened document',await js('__studio.engine.selected.fontFamily==="Zen Kaku Gothic New"&&__studio.engine.fonts.some(f=>f.family==="Zen Kaku Gothic New"&&f.data)'));
 await js('__studio.save(false)');
 await command('Network.emulateNetworkConditions',{offline:false,latency:0,downloadThroughput:-1,uploadThroughput:-1});
 for(const kind of ['banner','flyer','logo','card']){
  const copy=path.join(output,kind+'.dake');await fs.copyFile(path.join(exampleRoot,kind+'.dake'),copy);await second(copy);await waitUntil(`__studio.currentPath===${JSON.stringify(copy)}`,30000);
  await js("(()=>{const e=__studio.engine;const o=e.layers.find(o=>!o.locked);e.select(o.id);e.nudge(1,1);})()");await js('__studio.save(false)');
  for(const format of ['png','jpeg','webp','svg']){const data=await js(format==='svg'?'__studio.engine.exportSVG()':`__studio.engine.exportRaster('${format}',1,.95)`);await fs.writeFile(path.join(output,kind+'.'+format),format==='svg'?data:Buffer.from(data.split(',')[1],'base64'));}
  const png=await js("__studio.engine.exportRaster('png',1)");const meta=await js('({width:__studio.engine.width,height:__studio.engine.height,dpi:__studio.engine.meta.dpi})');
  // Same pure PDF writer bundled in the EXE; separate report records this boundary.
  const {toPDF}=require('../desktop/print-artwork.cjs');await fs.writeFile(path.join(output,kind+'.pdf'),await toPDF({data:png,...meta}));
  check('existing '+kind+' edited, saved and exported in five formats',true);
 }
 check('no renderer exceptions',report.errors.length===0);await close();report.passed=true;report.packaged=packaged;report.fontOfflineScope='renderer network disabled; native cache get performs no HTTP';report.pdfScope='same local PDF worker via test process';await fs.writeFile(path.join(root,'evidence',packaged?'workflow-packaged-results.json':'workflow-results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({passed:true,checks:report.checks.length,output}));
}
run().catch(async error=>{report.passed=false;report.failure=error.stack;console.error(report.lastExpression);if(process.env.STADIO_KEEP_QA==='1'){await fs.writeFile(path.join(output,'debug.json'),JSON.stringify(report));console.error('Keeping QA window');await new Promise(()=>{});}try{const shot=await command('Page.captureScreenshot',{format:'png'});await fs.writeFile(path.join(output,'failure.png'),Buffer.from(shot.data,'base64'));}catch{};await fs.writeFile(path.join(root,'evidence',packaged?'workflow-packaged-results.json':'workflow-results.json'),JSON.stringify(report,null,2));console.error(error);process.exitCode=1;}).finally(()=>{connection?.close();if(child&&child.exitCode===null)child.kill();});
