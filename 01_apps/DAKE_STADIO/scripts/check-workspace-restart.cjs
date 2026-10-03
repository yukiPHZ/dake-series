'use strict';
const path=require('node:path'),fs=require('node:fs'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');require('./local-test-env.cjs')();
const {app,BrowserWindow,dialog}=require('electron');
const fixture=process.env.STADIO_RESTART_FIXTURE;if(!fixture)throw Error('STADIO_RESTART_FIXTURE is required');
const data=JSON.parse(fs.readFileSync(fixture,'utf8')),source=data.canvas.objects[0].dakeComponent.link.path;
const sourceHash=crypto.createHash('sha256').update(fs.readFileSync(source)).digest('hex');
const output=path.join(root,'test-output','workspace','restart-'+Date.now());fs.mkdirSync(output,{recursive:true});process.env.STADIO_USER_DATA=path.join(output,'profile');
let openTarget=fixture;dialog.showOpenDialog=async()=>({canceled:false,filePaths:[openTarget]});dialog.showSaveDialog=async()=>({canceled:false,filePath:path.join(output,'reopened-edited-card.dake')});
require('../desktop/main.cjs');
let win;const report={pid:process.pid,fixture,output,checks:[],errors:[]},delay=ms=>new Promise(r=>setTimeout(r,ms));
const js=s=>win.webContents.executeJavaScript(s,true),check=(name,value)=>{assert.ok(value,name);report.checks.push({name,passed:true});};
async function wait(s){for(let n=0;n<600;n++){if(await js(s))return;await delay(25);}throw Error('wait '+s);}
(async()=>{
 await app.whenReady();win=BrowserWindow.getAllWindows()[0];win.webContents.on('console-message',e=>{if(e.level==='error')report.errors.push(e.message);});await wait('!!window.__studio?.workspace');
 await js('__studio.open()');check('fresh PID opens cached linked component without original capability',await js('__studio.engine.layers[0].dakeComponent.document.canvas.objects[0].text==="LINK UPDATED"'));
 await js('(async()=>{__studio.engine.select(__studio.engine.layers[0].id);await __studio.workspace.editComponent();__studio.engine.select(__studio.engine.layers[0].id);__studio.engine.updateSelected({text:"RESTART EDIT"});await __studio.workspace.finishChild(true);return true;})()');
 check('fresh process can edit child text without flattening',await js('__studio.engine.layers[0].dakeComponent.document.canvas.objects[0].text==="RESTART EDIT"&&__studio.engine.layers[0].getObjects().some(o=>o.text==="RESTART EDIT")'));
 openTarget=source;await js('window.refreshDone=false;window.refreshError=null;__studio.engine.select(__studio.engine.layers[0].id);__studio.workspace.refreshComponent().then(()=>window.refreshDone=true).catch(e=>{window.refreshError=e.message;window.refreshDone=true});true');
 await wait('!!document.getElementById("dialog-select")');check('expired capability asks before reading external source',true);await js('document.getElementById("dialog-select").click();true');await wait('window.refreshDone');
 check('explicit source reauthorization works after restart',await js('!window.refreshError&&__studio.engine.layers[0].dakeComponent.document.canvas.objects[0].text==="LINK UPDATED"'));
 await js('__studio.engine.undo()');check('source refresh undo restores unsaved child edit',await js('__studio.engine.layers[0].dakeComponent.document.canvas.objects[0].text==="RESTART EDIT"'));
 check('fresh process saves reeditable output',await js('__studio.save(true)'));
 const saved=JSON.parse(fs.readFileSync(path.join(output,'reopened-edited-card.dake'),'utf8'));check('second save retains linked editable child text',saved.canvas.objects[0].dakeComponent.document.canvas.objects[0].text==='RESTART EDIT');
 check('original linked file remains byte-identical',sourceHash===crypto.createHash('sha256').update(fs.readFileSync(source)).digest('hex'));
 fs.writeFileSync(path.join(output,'reopened.png'),(await win.webContents.capturePage()).toPNG());check('no renderer errors',report.errors.length===0);report.passed=true;fs.writeFileSync(path.join(output,'results.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report));app.exit(0);
})().catch(error=>{report.error=error.stack;fs.writeFileSync(path.join(output,'results.json'),JSON.stringify(report,null,2));console.error(error);app.exit(1);});
