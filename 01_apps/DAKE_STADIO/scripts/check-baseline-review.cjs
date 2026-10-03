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
const executable = path.resolve(process.argv[2] || path.join(root, 'dist', 'DAKE_STADIO-win32-x64', 'DAKE_STADIO.exe'));
const output = path.join(root, 'test-output', `baseline-review-${Date.now()}`);
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
  const result = await command('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
}
async function waitUntil(expression, timeout = 15000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) { if (await js(expression)) return; await delay(50); }
  throw new Error(`Packaged UI timeout: ${expression}`);
}

async function run(){
 await fs.mkdir(output,{recursive:true});const port=await freePort();const env={...process.env,STADIO_USER_DATA:path.join(output,'profile')};delete env.ELECTRON_RUN_AS_NODE;
 const project=path.join(root,'artifacts','baseline-0.2.0','examples','card.dake');child=spawn(executable,['--remote-debugging-port='+port,project],{cwd:output,env,stdio:'ignore',windowsHide:true});
 let target;for(let i=0;i<100;i++){try{target=(await(await fetch('http://127.0.0.1:'+port+'/json/list')).json()).find(t=>t.type==='page');if(target)break;}catch{}await delay(100);}
 connection=new WebSocket(target.webSocketDebuggerUrl);await new Promise(r=>connection.addEventListener('open',r,{once:true}));connection.addEventListener('message',event=>{const m=JSON.parse(event.data);if(m.id&&pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);clearTimeout(p.timer);m.error?p.reject(new Error(JSON.stringify(m.error))):p.resolve(m.result);}});await waitUntil('!!window.__studio');
 check('0.2 ignores existing project path on cold launch',await js('__studio.currentPath===null&&__studio.engine.layers.length===0'));
 await js("document.querySelector('[data-tool=text]').click();document.querySelector('#font-library').click()");await waitUntil('!!document.querySelector("#font-google")');await js("document.querySelector('#font-google').click()");await waitUntil('document.querySelector("#font-google").classList.contains("active")');
 await js("document.querySelector('#font-search').value='Zen Kaku Gothic New';document.querySelector('#font-search').dispatchEvent(new Event('input'));document.querySelector('.font-row button').click()");await waitUntil('__studio.engine.selected?.fontFamily==="Zen Kaku Gothic New"',120000);
 check('0.2 actually acquires and applies Google family',await js('__studio.engine.fonts.some(f=>f.family==="Zen Kaku Gothic New"&&f.data)'));await js("document.querySelector('#dialog-close').click();void __studio.production.newDocument()");await waitUntil('!!document.querySelector("#dialog-discard")');await js("document.querySelector('#dialog-discard').click()");await waitUntil('!!document.querySelector("#dialog-create")');await js("document.querySelector('#dialog-create').click()");await waitUntil('__studio.engine.layers.length===0');await js("document.querySelector('[data-tool=text]').click()");
 check('0.2 new document loses downloaded font choice and active Face',await js("![...document.querySelectorAll('#available-fonts option')].some(o=>o.value==='Zen Kaku Gothic New')&&![...document.fonts].some(f=>f.family==='Zen Kaku Gothic New')"));
 await js("document.querySelector('.layer-row').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}))");await waitUntil('!!document.querySelector("#rename-value")');check('0.2 layer double-click handler opens rename dialog',true);
 await fs.writeFile(path.join(root,'evidence','baseline-review-results.json'),JSON.stringify({...report,passed:true},null,2));await js('stadio.closeReady()');console.log(JSON.stringify({passed:true,checks:report.checks.length,output}));
}
run().catch(e=>{console.error(e);process.exitCode=1}).finally(()=>{connection?.close();if(child&&child.exitCode===null)child.kill()});
