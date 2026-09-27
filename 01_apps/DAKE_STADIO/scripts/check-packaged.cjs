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
const output = path.join(root, 'test-output', `packaged-smoke-${Date.now()}`);
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
async function run() {
  await fs.mkdir(output, { recursive: true });
  const port = await freePort();
  const environment = { ...process.env, STADIO_USER_DATA: path.join(output, 'profile') };
  delete environment.ELECTRON_RUN_AS_NODE;
  const start = Date.now();
  child = spawn(executable, [`--remote-debugging-port=${port}`, '--remote-debugging-address=127.0.0.1'], { cwd: output, env: environment, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
  child.stdout.resume(); child.stderr.resume();
  let target;
  const deadline = Date.now() + 20000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) throw new Error(`Packaged app exited before ready: ${child.exitCode}`);
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
      target = targets.find(item => item.type === 'page' && item.url.includes('app.asar/build/index.html'));
      if (target) break;
    } catch {}
    await delay(100);
  }
  if (!target) throw new Error('No packaged page exposed by local test-only debug port');
  report.pageUrl = target.url;
  check('executable loads page from packaged app.asar', target.url.includes('resources/app.asar/build/index.html'));
  connection = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { connection.addEventListener('open', resolve, { once: true }); connection.addEventListener('error', reject, { once: true }); });
  connection.addEventListener('message', event => {
    const message = JSON.parse(event.data);
    if (message.id && pending.has(message.id)) {
      const item = pending.get(message.id); pending.delete(message.id); clearTimeout(item.timer);
      message.error ? item.reject(new Error(JSON.stringify(message.error))) : item.resolve(message.result);
    } else if (message.method === 'Runtime.exceptionThrown') report.errors.push(message.params.exceptionDetails);
  });
  await command('Runtime.enable');
  await command('Page.enable');
  await waitUntil('!!window.__studio && !!window.stadio');
  report.readyMs = Date.now() - start;
  check('packaged Japanese interface and native bridge initialize independently', await js(`document.getElementById('save').textContent === '保存' && __studio.engine.layers.length===0`));
  await js(`document.getElementById('sample').click()`);
  await waitUntil('__studio.engine.layers.length===8');
  check('packaged sample combines raster and editable vector layers', await js(`__studio.engine.layers.some(o=>o.type.toLowerCase()==='image') && __studio.engine.layers.some(o=>o.type.toLowerCase()==='path')`));
  const worker = await js(`(async()=>{const image=__studio.engine.layers.find(o=>o.type.toLowerCase()==='image');__studio.engine.select(image.id);await __studio.engine.applyImageAdjustment({brightness:.1});return {active:!!__studio.engine._worker,brightness:__studio.engine.selected.adjustments.brightness};})()`);
  check('packaged raster worker actually processes local pixels', worker.active && worker.brightness === .1);
  const document = await js('__studio.engine.serialize()');
  await js('window.stadio.autosave(__studio.engine.serialize())');
  const recovery = JSON.parse(await fs.readFile(path.join(output, 'profile', 'recovery.json'), 'utf8'));
  check('packaged native bridge writes recoverable local data', recovery.data.canvas.objects.length === 8 && recovery.data.format === 'dake-stadio');
  await fs.writeFile(path.join(output, 'packaged-render.dake'), JSON.stringify(document));
  const png = await js(`__studio.engine.exportRaster('png',1)`);
  const bytes = Buffer.from(png.split(',')[1], 'base64');
  check('packaged engine exports correct-size PNG', bytes.readUInt32BE(16) === 1200 && bytes.readUInt32BE(20) === 800);
  await fs.writeFile(path.join(output, 'packaged-render.png'), bytes);
  const screenshot = await command('Page.captureScreenshot', { format: 'png' });
  await fs.writeFile(path.join(root, 'evidence', 'packaged-screen.png'), Buffer.from(screenshot.data, 'base64'));
  check('portable distribution contains editable example and required licenses', (await fs.stat(path.join(path.dirname(executable), 'examples', 'はじめの作品.dake'))).size > 100 && (await fs.stat(path.join(path.dirname(executable), 'LICENSES.chromium.html'))).size > 100 && (await fs.stat(path.join(path.dirname(executable), 'THIRD_PARTY_NOTICES.txt'))).size > 100);
  const modern = await js('__studio.engine.serialize().version===2');
  if (modern) {
    const catalog = await js('window.stadio.listGoogleFonts()');
    check('packaged offline Google catalog includes Japanese families', catalog.families.length >= 1900 && catalog.families.filter(item=>item.subsets.includes('japanese')).length >= 60);
    const local = await js('window.stadio.listLocalFonts()');
    check('packaged trusted preload discovers installed Windows fonts', local.length > 20 && local.some(font=>font.postscriptName==='Meiryo'));
    const vector = await js("(async()=>{const e=__studio.engine;const a=e.addShape('rect',{left:30,top:30,width:80,height:80});const b=e.addShape('ellipse',{left:65,top:40,rx:45,ry:40});e.selectMany([a.id,b.id]);const result=await e.booleanSelected('union');return {worker:!!e._vectorWorker,paths:result.length,type:result[0]?.type.toLowerCase()}})()");
    check('packaged vector worker computes editable Boolean paths', vector.worker && vector.paths > 0 && vector.type==='path');
    for(const kind of ['banner','flyer','logo','card'])for(const ext of ['dake','png','svg','pdf'])check('portable '+kind+' '+ext+' example exists',(await fs.stat(path.join(path.dirname(executable),'examples',kind+'.'+ext))).size>100);
  }
  check('packaged renderer has no uncaught exceptions', report.errors.length === 0);
  await js('window.stadio.discardRecovery()');
  const exited = new Promise(resolve => child.once('exit', code => resolve(code)));
  await js('window.stadio.closeReady()');
  const exitCode = await Promise.race([exited, delay(5000).then(() => 'timeout')]);
  check('packaged application exits successfully', exitCode === 0);
  report.runtimeSha256 = {
    executable: await hash(executable),
    appAsar: await hash(path.join(path.dirname(executable), 'resources', 'app.asar'))
  };
  report.passed = true;
  await fs.writeFile(path.join(root, 'evidence', 'packaged-results.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}
run().catch(async error => {
  report.passed = false; report.failure = error.stack;
  await fs.writeFile(path.join(root, 'evidence', 'packaged-results.json'), JSON.stringify(report, null, 2));
  console.error(error); process.exitCode = 1;
}).finally(() => {
  connection?.close();
  if (child && child.exitCode === null) child.kill();
});
