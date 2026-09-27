require('./local-test-env.cjs')();
'use strict';
const {app,BrowserWindow,session,ipcMain}=require('electron');
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const {createFontService}=require('../desktop/fonts.cjs');
const {checksum}=require('../desktop/font-binary.cjs');
const opentype=require('opentype.js');
const root=path.resolve(__dirname,'..'),out=path.join(root,'test-output','font-probe-v2');
fs.mkdirSync(out,{recursive:true});app.setPath('userData',path.join(out,'profile'));
const html=path.join(out,'index.html');fs.writeFileSync(html,'<!doctype html><html><meta charset="UTF-8"><body>Font probe</body></html>');
const service=createFontService({userData:app.getPath('userData'),catalogPath:path.join(root,'assets','google-fonts-catalog.json')});
ipcMain.handle('stadio:prepareLocalFont',(_e,request)=>service.prepareLocal(request));
const report={checks:[],local:[]};const check=(name,ok)=>{assert.ok(ok,name);report.checks.push(name)};
app.whenReady().then(async()=>{
 session.defaultSession.setPermissionCheckHandler((_w,p)=>p==='local-fonts');
 session.defaultSession.setPermissionRequestHandler((_w,p,c)=>c(p==='local-fonts'));
 const w=new BrowserWindow({show:false,webPreferences:{preload:path.join(root,'desktop','preload.cjs'),contextIsolation:true,sandbox:true,nodeIntegration:false}});
 await w.loadFile(html);
 const fonts=await w.webContents.executeJavaScript('window.stadio.listLocalFonts()',true);
 report.localCount=fonts.length;check('isolated preload enumerates installed Windows font faces',fonts.length>20);
 for(const name of ['BIZ-UDGothic','Meiryo-Bold']){
  const descriptor=fonts.find(f=>f.postscriptName===name);assert.ok(descriptor);
  const result=await w.webContents.executeJavaScript('window.stadio.getLocalFont('+JSON.stringify(descriptor.id)+')',true);
  const bytes=Buffer.from(result.data.split(',')[1],'base64');
  check('TTC face '+name+' extracts to single SFNT',bytes.toString('ascii',0,4)!=='ttcf');
  check('extracted SFNT checksum valid for '+name,checksum(bytes)===0xb1b0afba);
  const font=opentype.parse(bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength));
  check('selected Windows face matches requested PostScript name '+name,JSON.stringify(font.names).includes(name));
  check('Japanese outlines exist for '+name,font.getPath('日本語',0,80,72).commands.length>40);
  report.local.push({family:result.family,postscriptName:result.postscriptName,weight:result.weight,fsType:result.fsType,sha256:result.sha256,bytes:bytes.length});
 }
 const catalog=await service.list();report.catalogCount=catalog.families.length;check('offline bundled Google catalog has Japanese families',catalog.families.filter(f=>f.subsets.includes('japanese')).length>=60);
 const google=await service.download({family:'Noto Sans JP',id:'ofl/notosansjp',requestId:'qa-notosansjp'});
 check('Google font download includes full license and hash',google.license.includes('SIL OPEN FONT LICENSE')&&google.sha256.length===64);
 const bytes=Buffer.from(google.data.split(',')[1],'base64');
 fs.writeFileSync(path.join(out,'NotoSansJP.ttf'),bytes);fs.writeFileSync(path.join(out,'NotoSansJP-OFL.txt'),google.license);
 const second=await service.download({family:'Noto Sans JP',id:'ofl/notosansjp'});check('second retrieval comes from verified local cache',second.cached&&second.sha256===google.sha256);
 const font=opentype.parse(bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength));
 check('downloaded full font supports Japanese paths',font.getPath('日本語の制作',0,80,72).commands.length>40);
 report.google={family:google.family,weight:google.weight,weightRange:google.weightRange,sha256:google.sha256,sourceSha256:google.sourceSha256,fsType:google.fsType,bytes:bytes.length,fixture:path.join(out,'NotoSansJP.ttf'),license:path.join(out,'NotoSansJP-OFL.txt')};
 report.passed=true;fs.writeFileSync(path.join(root,'evidence','native-font-results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));app.exit(0);
}).catch(error=>{report.passed=false;report.failure=error.stack;fs.writeFileSync(path.join(root,'evidence','native-font-results.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1)});


