require('./local-test-env.cjs')();
'use strict';
const {app,BrowserWindow}=require('electron');
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..'),out=path.join(root,'test-output','native-pdf-v2');
fs.mkdirSync(out,{recursive:true});app.setPath('userData',path.join(out,'profile'));
const {toPDF,dimensions}=require('../desktop/print-artwork.cjs');
const report={checks:[],pdfs:[],physicalPrinter:'not tested; no print job was submitted'};
const check=(name,ok)=>{assert.ok(ok,name);report.checks.push(name)};
app.whenReady().then(async()=>{
 const draw=new BrowserWindow({show:false,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}});
 await draw.loadURL('data:text/html,<html><body></body></html>');
 for(const spec of [{name:'business-card-bleed',width:1146,height:720,dpi:300},{name:'a5-landscape',width:1748,height:1240,dpi:212},{name:'transparent-composite',width:800,height:500,dpi:144,transparent:true}]){
  const data=await draw.webContents.executeJavaScript(`(()=>{const c=document.createElement('canvas');c.width=${spec.width};c.height=${spec.height};const x=c.getContext('2d');x.fillStyle='#ffffff';if(!${!!spec.transparent})x.fillRect(0,0,c.width,c.height);x.fillStyle=${JSON.stringify(spec.transparent?'rgba(179,51,102,0.5)':'#b53366')};x.fillRect(0,0,c.width/2,c.height/2);x.fillStyle='#126779';x.fillRect(c.width/2,c.height/2,c.width/2,c.height/2);x.strokeStyle='#000000';x.lineWidth=12;x.strokeRect(6,6,c.width-12,c.height-12);return c.toDataURL('image/png')})()`);
  fs.writeFileSync(path.join(out,spec.name+'.png'),Buffer.from(data.split(',')[1],'base64'));
  const started=performance.now(),bytes=await toPDF({...spec,data});
  const text=bytes.toString('latin1');const media=/\/MediaBox\s*\[\s*([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*\]/.exec(text);
  check(spec.name+' produces a real PDF',bytes.toString('ascii',0,5)==='%PDF-');
  check(spec.name+' is one page',(text.match(/\/Type\s*\/Page\b/g)||[]).length===1);
  check(spec.name+' has physical page dimensions within 0.00001 point',media&&Math.abs(Number(media[3])-spec.width/spec.dpi*72)<.00001&&Math.abs(Number(media[4])-spec.height/spec.dpi*72)<.3);
  check(spec.name+' contains the raster artwork',/\/Subtype\s*\/Image\b/.test(text));
  fs.writeFileSync(path.join(out,spec.name+'.pdf'),bytes);
  report.pdfs.push({...spec,mediaBox:media.slice(1).map(Number),bytes:bytes.length,elapsedMs:Math.round(performance.now()-started)});
  assert.throws(()=>dimensions({...spec,data,width:spec.width-1}));check(spec.name+' rejects incorrect image dimensions',true);
 }
 draw.destroy();
 report.passed=true;fs.writeFileSync(path.join(root,'evidence','native-pdf-results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));app.exit(0);
}).catch(error=>{report.passed=false;report.failure=error.stack;fs.writeFileSync(path.join(root,'evidence','native-pdf-results.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1)});
