'use strict';
require('./local-test-env.cjs')();
const path=require('node:path'),fs=require('node:fs'),root=path.resolve(__dirname,'..'),out=path.join(root,'test-output','imaging-v4');
if(!process.versions.electron){
 fs.mkdirSync(out,{recursive:true});
 require('esbuild').buildSync({entryPoints:[path.join(root,'tests/imaging-browser.js')],outfile:path.join(out,'imaging.bundle.js'),bundle:true,format:'iife',platform:'browser',logLevel:'silent'});
 fs.copyFileSync(path.join(root,'src/raster-worker.js'),path.join(out,'raster-worker.js'));
 fs.writeFileSync(path.join(out,'index.html'),'<!doctype html><meta charset="utf-8"><title>DAKE imaging verification</title><canvas id="editor"></canvas><script src="imaging.bundle.js"></script>');
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
 const child=require('node:child_process').spawnSync(require('electron'),[__filename],{cwd:root,env,stdio:'inherit',windowsHide:true,timeout:180000});process.exit(child.status??1);
}else{
 const {app,BrowserWindow}=require('electron');app.setPath('userData',path.join(out,'profile'));
 app.whenReady().then(async()=>{
  const w=new BrowserWindow({show:false,width:1000,height:800,webPreferences:{contextIsolation:true,nodeIntegration:false,sandbox:true,backgroundThrottling:false}});
  try{
   await w.loadFile(path.join(out,'index.html'));const report=await w.webContents.executeJavaScript('window.runImagingChecks()');
   fs.writeFileSync(path.join(out,'recipe-mask.dake'),JSON.stringify(report.save));delete report.save;
   fs.writeFileSync(path.join(out,'recipe-mask.png'),Buffer.from(report.png.split(',')[1],'base64'));delete report.png;
   w.webContents.sendInputEvent({type:'mouseMove',...report.native});w.webContents.sendInputEvent({type:'mouseDown',...report.native,button:'left',clickCount:1});
   await new Promise(r=>setTimeout(r,50));w.webContents.sendInputEvent({type:'mouseUp',...report.native,button:'left',clickCount:1});await new Promise(r=>setTimeout(r,100));
   const native=await w.webContents.executeJavaScript('window.finishNativeMaskCheck()');report.results.push({name:'native Electron pointer paints mask only and commits one undo',...native});report.passed=report.results.every(r=>r.passed);
   fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));app.exit(report.passed?0:1);
  }catch(error){console.error(error);fs.writeFileSync(path.join(out,'failure.txt'),String(error.stack||error));app.exit(1);}
 });
}

