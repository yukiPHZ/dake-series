require('./local-test-env.cjs')();
'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { spawn } = require('node:child_process');
const assert = require('node:assert/strict');
const electron = require('electron');
const root = path.resolve(__dirname, '..');
const output = path.join(root, 'test-output', `restart-${Date.now()}`);
async function phase(name) {
  const environment = { ...process.env, STADIO_RESTART_OUTPUT: output };
  delete environment.ELECTRON_RUN_AS_NODE;
  await new Promise((resolve, reject) => {
    const child = spawn(electron, ['scripts/check-restart.cjs', name], { cwd: root, env: environment, stdio: 'inherit', windowsHide: true });
    const timeout = setTimeout(() => { child.kill(); reject(new Error(`${name} timed out after 40 seconds`)); }, 40000);
    child.once('error', error => { clearTimeout(timeout); reject(error); });
    child.once('exit', code => { clearTimeout(timeout); code === 0 ? resolve() : reject(new Error(`${name} exited with ${code}`)); });
  });
  return JSON.parse(await fs.readFile(path.join(output, `${name}.json`), 'utf8'));
}
async function run() {
  await fs.mkdir(output, { recursive: true });
  const prepare = await phase('prepare');
  const resume = await phase('resume');
  assert.notEqual(prepare.processId, resume.processId);
  const report = { passed: prepare.passed && resume.passed, output, processes: [prepare.processId, resume.processId], checks: [...prepare.checks, ...resume.checks, 'prepare and resume ran in distinct Electron processes'], interruption: prepare.terminationMode };
  await fs.writeFile(path.join(root, 'evidence', 'restart-results.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}
run().catch(error => { console.error(error); process.exitCode = 1; });
