'use strict';
require('./local-test-env.cjs')();
const fs=require('node:fs'),path=require('node:path'),root=path.resolve(__dirname,'..'),out=path.join(root,'test-output','imaging-ui-v4');
if(!process.versions.electron){
 fs.mkdirSync(out,{recursive:true});const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
 require('./build.cjs')().then(()=>{const p=require('node:child_process').spawnSync(require('electron'),[__filename],{cwd:root,env,stdio:'inherit',windowsHide:true,timeout:180000});process.exit(p.status??1);});
}else{
 const {app,BrowserWindow,dialog}=require('electron');process.env.STADIO_USER_DATA=path.join(out,'profile-'+Date.now());
 let savePath=path.join(out,'corrected-mask.dake');dialog.showSaveDialog=async()=>({canceled:false,filePath:savePath});dialog.showOpenDialog=async()=>({canceled:true,filePaths:[]});
 require('../desktop/main.cjs');let win;const report={checks:[],errors:[],method:'Native Electron pointer button/mask gestures; fixture insertion and slider input/change events through renderer; native save dialog path supplied by harness.'};
 const delay=ms=>new Promise(r=>setTimeout(r,ms)),js=s=>win.webContents.executeJavaScript(s,true),check=(name,ok)=>{report.checks.push({name,passed:!!ok});if(!ok)throw new Error(name);};
 const wait=async(exp,ms=12000)=>{const end=Date.now()+ms;while(Date.now()<end){if(await js(exp))return;await delay(30);}throw new Error('Timeout '+exp);};
 async function click(sel){const p=await js('(()=>{const e=document.querySelector('+JSON.stringify(sel)+');if(!e)throw new Error("missing element");e.scrollIntoView({block:"nearest"});const r=e.getBoundingClientRect();return{x:Math.round(r.left+r.width/2),y:Math.round(r.top+r.height/2)};})()');win.webContents.sendInputEvent({type:'mouseDown',...p,button:'left',clickCount:1});win.webContents.sendInputEvent({type:'mouseUp',...p,button:'left',clickCount:1});await delay(70);}
 async function input(id,value,event='input'){await js('(()=>{const e=document.getElementById('+JSON.stringify(id)+');e.value='+JSON.stringify(value)+';e.dispatchEvent(new Event('+JSON.stringify(event)+',{bubbles:true}));})()');await delay(40);}
 const screen=async name=>{win.showInactive();await delay(160);const b=win.getContentBounds();fs.writeFileSync(path.join(out,name),(await win.webContents.capturePage({x:0,y:0,width:b.width,height:b.height},{stayAwake:true})).toPNG());win.hide();};
 async function flow(){
  await app.whenReady();win=BrowserWindow.getAllWindows()[0];win.hide();win.webContents.on('console-message',e=>{if(e.level==='error')report.errors.push(e.message);});
  if(win.webContents.isLoading())await new Promise(r=>win.webContents.once('did-finish-load',r));await wait('!!window.__studio?.imaging');
  await js('(async()=>{const e=__studio.engine;await e.newDocument({width:800,height:500,name:"補正とマスクの操作検証",background:"#fff"});const c=document.createElement("canvas");c.width=640;c.height=400;const g=c.getContext("2d");const gradient=g.createLinearGradient(0,0,640,400);gradient.addColorStop(0,"#4b798c");gradient.addColorStop(1,"#d19951");g.fillStyle=gradient;g.fillRect(0,0,640,400);g.fillStyle="#eedab1";g.fillRect(220,60,180,240);const im=await e.importImage(c.toDataURL(),"generated-fixture");e.updateSelected({left:80,top:50,scaleX:1,scaleY:1});window.qaImageId=im.id;window.qaImageSource=im.toObject().src;window.qaBefore=await e.exportRaster();e.fit(850,650);return true;})()');
  await click('[data-panel="effects"]');check('effects has compact correction entry',await js('!!document.querySelector("#image-adjust-open")&&!document.querySelector("#adjust-brightness")'));
  await click('#image-adjust-open');await wait('!!document.querySelector("#im-adjust-preview")');await input('im-adjust-exposure',1);await wait('__studio.engine.selected.dakeImageEdit?.exposure===1&&!__studio.engine._workerJob');
  check('dialog previews actual corrected pixels and retains source',await js('__studio.engine.selected.toObject().src===qaImageSource&&document.querySelector("#im-adjust-preview").getContext("2d").getImageData(300,80,1,1).data[3]>0'));
  await screen('01-adjust-live.png');await click('#dialog-cancel');
  check('cancel restores original canvas',await js('__studio.engine.exportRaster().then(p=>p===qaBefore)'));
  await click('#image-adjust-open');await input('im-adjust-exposure',.8);await click('[data-im-page="color"]');await input('im-adjust-vibrance',.4);await click('[data-im-page="curve"]');await screen('02-curve.png');await click('#dialog-apply');await wait('!__studio.engine.imagingPreviewing');
  check('UI apply retains exposure and natural saturation recipe',await js('__studio.engine.selected.dakeImageEdit.exposure===.8&&__studio.engine.selected.dakeImageEdit.vibrance===.4'));
  await click('[data-im-target="layerMask"]');await wait('__studio.engine._imageEditingTarget==="layerMask"&&!__studio.engine.busy');
  await input('im-mask-hardness',100);await input('im-mask-size',90);
  const p=await js('(()=>{const e=__studio.engine,r=e.canvas.upperCanvasEl.getBoundingClientRect(),v=e.canvas.viewportTransform;return{x:Math.round(r.left+v[4]+400*v[0]),y:Math.round(r.top+v[5]+250*v[3])}})()');
  const history=await js('__studio.engine._historyIndex');
  win.webContents.sendInputEvent({type:'mouseDown',...p,button:'left',clickCount:1});await delay(70);win.webContents.sendInputEvent({type:'mouseUp',...p,button:'left',clickCount:1});await wait('!__studio.engine.busy');
  check('native mask brush keeps source and creates one undo',await js('__studio.engine._historyIndex=== '+(history+1)+' &&__studio.engine.selected.toObject().src===qaImageSource'));
  await screen('03-mask-painted.png');await click('#im-mask-restore');win.webContents.sendInputEvent({type:'mouseDown',...p,button:'left',clickCount:1});win.webContents.sendInputEvent({type:'mouseUp',...p,button:'left',clickCount:1});await wait('!__studio.engine.busy');
  await click('#im-mask-select');const a={x:p.x-90,y:p.y-75},b={x:p.x+100,y:p.y+85};win.webContents.sendInputEvent({type:'mouseDown',...a,button:'left',clickCount:1});win.webContents.sendInputEvent({type:'mouseMove',...b,modifiers:['leftButtonDown']});win.webContents.sendInputEvent({type:'mouseUp',...b,button:'left',clickCount:1});await wait('!__studio.engine.busy&&!__studio.engine._strokeTask');
  check('rectangle selection mask available through UI',await js('!!__studio.engine.selected.dakeLayerMask?.src'));
  await input('im-mask-feather',10);await input('im-mask-feather',10,'change');await wait('!__studio.engine.imagingPreviewing&&!__studio.engine._workerJob');
  await screen('04-mask-selection-feather.png');await click('#save');await wait('!__studio.dirty');
  const saved=JSON.parse(fs.readFileSync(savePath,'utf8'));check('native save retains original plus recipe and mask',saved.canvas.objects.some(o=>o.src&&o.dakeImageEdit?.exposure===.8&&o.dakeLayerMask?.feather===10));
  check('UI integration has no unexpected renderer exceptions',report.errors.length===0);report.passed=true;
  fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));app.exit(0);
 }
 flow().catch(async error=>{report.passed=false;report.error=String(error.stack||error);try{await screen('failure.png');report.dom=await js('document.body.innerText');}catch{}fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1);});
}

