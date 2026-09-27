'use strict';
require('./local-test-env.cjs')();const {app,BrowserWindow}=require('electron');const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),output=path.join(root,'test-output','pdf-performance-'+Date.now());fs.mkdirSync(output,{recursive:true});app.setPath('userData',path.join(output,'profile'));
const {toPDF}=require('../desktop/print-artwork.cjs');
app.whenReady().then(async()=>{
 const win=new BrowserWindow({show:false,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});await win.loadURL('data:text/html,<html></html>');
 const width=8000,height=4000,dpi=300;
 const data=await win.webContents.executeJavaScript(`(()=>{const c=document.createElement('canvas');c.width=8000;c.height=4000;const x=c.getContext('2d');const g=x.createLinearGradient(0,0,8000,4000);g.addColorStop(0,'#193a6a');g.addColorStop(.5,'#f5ce8c');g.addColorStop(1,'#742453');x.fillStyle=g;x.fillRect(0,0,8000,4000);return c.toDataURL()})()`);
 let previous=performance.now(),maximum=0,ticks=0;const tick=setInterval(()=>{const now=performance.now();maximum=Math.max(maximum,now-previous);previous=now;ticks++},10);
 const started=performance.now();let bytes;try{bytes=await toPDF({data,width,height,dpi})}finally{clearInterval(tick)}
 assert.ok(bytes.toString('ascii',0,5)==='%PDF-');assert.ok(ticks>1&&maximum<500);
 fs.writeFileSync(path.join(output,'32-megapixel.pdf'),bytes);const report={passed:true,width,height,elapsedMs:Math.round(performance.now()-started),maxMainEventGapMs:Math.round(maximum),responsiveTicks:ticks,pdfBytes:bytes.length,output,checks:['32-megapixel PDF completed','main event loop continued with maximum gap under 500ms']};
 fs.writeFileSync(path.join(root,'evidence','native-pdf-performance.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));win.destroy();app.exit(0);
}).catch(error=>{console.error(error);app.exit(1)});
