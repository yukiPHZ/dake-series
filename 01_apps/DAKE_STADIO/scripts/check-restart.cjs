require('./local-test-env.cjs')();
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { app, BrowserWindow, dialog } = require('electron');
const root = path.resolve(__dirname, '..');
const output = process.env.STADIO_RESTART_OUTPUT;
const phase = process.argv[2];
if (!output || !['prepare', 'resume'].includes(phase)) throw new Error('Expected isolated restart output and prepare/resume phase');
process.env.STADIO_USER_DATA = path.join(output, 'profile');
const projectPath = path.join(output, '再起動検証.dake');
dialog.showSaveDialog = async () => ({ canceled: false, filePath: projectPath });
dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [projectPath] });
dialog.showErrorBox = (title, text) => { throw new Error(`${title}: ${text}`); };
require('../desktop/main.cjs');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const report = { phase, processId: process.pid, started: new Date().toISOString(), checks: [] };
let win;
const js = script => win.webContents.executeJavaScript(script, true);
const check = (name, value) => { assert.ok(value, name); report.checks.push(name); };
const click = async id => { await js(`document.getElementById(${JSON.stringify(id)}).click()`); await delay(60); };
async function waitUntil(script, timeout = 12000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) { if (await js(script)) return; await delay(30); }
  throw new Error(`Timed out: ${script}`);
}
async function run() {
  await app.whenReady();
  win = BrowserWindow.getAllWindows()[0];
  if (win.webContents.isLoading()) await new Promise(resolve => win.webContents.once('did-finish-load', resolve));
  await waitUntil('!!window.__studio');
  if (phase === 'prepare') {
    await click('sample'); await waitUntil('__studio.engine.layers.length === 8');
    await js(`__studio.engine.renameDocument('再起動前に保存した作品')`);
    await click('save'); await waitUntil('!__studio.dirty');
    check('first process saves editable project with raster and vectors', JSON.parse(fs.readFileSync(projectPath, 'utf8')).canvas.objects.length === 8);
    await js(`__studio.engine.renameDocument('自動復旧で戻る作品');void __studio.engine.addShape('rect')`);
    await waitUntil('__studio.dirty && __studio.engine.layers.length === 9');
    await delay(2300);
    const recovery = JSON.parse(fs.readFileSync(path.join(process.env.STADIO_USER_DATA, 'recovery.json'), 'utf8'));
    check('first process autosaves unsaved name and added layer', recovery.data.name === '自動復旧で戻る作品' && recovery.data.canvas.objects.length === 9);
    report.terminationMode = 'app.exit bypasses window close cleanup, simulating interruption after autosave';
  } else {
    await waitUntil(`!!document.getElementById('recover')`);
    check('new process presents recovery UI for previous process data', true);
    await click('recover'); await waitUntil('__studio.dirty && __studio.engine.layers.length === 9');
    check('recovery UI restores previous unsaved name and layer graph', await js(`__studio.engine.name === '自動復旧で戻る作品'`));
    check('recovered content still contains editable vector and embedded image', await js(`__studio.engine.layers.some(o=>o.type.toLowerCase()==='path') && __studio.engine.layers.some(o=>o.type.toLowerCase()==='image')`));
    await click('open'); await waitUntil(`!!document.getElementById('dialog-discard')`); await click('dialog-discard');
    await waitUntil(`!__studio.dirty && __studio.engine.name === '再起動前に保存した作品'`);
    check('new process reopens saved file independently of old process', await js('__studio.engine.layers.length === 8'));
    const change = await js(`(()=>{const target=__studio.engine.layers.find(o=>o.type.toLowerCase()==='path');__studio.engine.select(target.id);const left=target.left+25;__studio.engine.updateSelected({left});return {id:target.id,left};})()`);
    check('reopened vector remains editable and marks document dirty', await js('__studio.dirty'));
    await click('save'); await waitUntil('!__studio.dirty');
    const saved = JSON.parse(fs.readFileSync(projectPath, 'utf8'));
    check('post-restart edit saves as editable path to same project', saved.canvas.objects.find(o => o.id === change.id).left === change.left);
    check('post-restart save preserves previous project as backup', fs.existsSync(`${projectPath}.bak`));
    check('clean save removes obsolete recovery from earlier process', !fs.existsSync(path.join(process.env.STADIO_USER_DATA, 'recovery.json')));
  }
  report.passed = true;
  fs.writeFileSync(path.join(output, `${phase}.json`), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
  if (phase === 'prepare') app.exit(0);
  else win.close();
}
run().catch(error => {
  report.passed = false; report.failure = error.stack;
  fs.writeFileSync(path.join(output, `${phase}.json`), JSON.stringify(report, null, 2));
  console.error(error); app.exit(1);
});
