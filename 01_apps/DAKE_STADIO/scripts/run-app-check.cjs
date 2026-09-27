require('./local-test-env.cjs')();
'use strict';
const path = require('node:path');
const { spawn } = require('node:child_process');
const electron = require('electron');
const root = path.resolve(__dirname, '..');
const environment = { ...process.env };
delete environment.ELECTRON_RUN_AS_NODE;
const child = spawn(electron, ['scripts/check-app.cjs'], {
  cwd: root, stdio: 'inherit', windowsHide: true,
  env: environment
});
const timeout = setTimeout(() => {
  console.error('App flow check exceeded 120 seconds');
  child.kill();
}, 120000);
child.once('error', error => { clearTimeout(timeout); console.error(error); process.exitCode = 1; });
child.once('exit', code => { clearTimeout(timeout); process.exitCode = code ?? 1; });
