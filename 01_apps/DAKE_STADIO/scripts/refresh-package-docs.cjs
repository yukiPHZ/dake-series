require('./local-test-env.cjs')();
'use strict';
const path = require('node:path');
const fs = require('node:fs/promises');
const { spawnSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const version = require('../package.json').version;
const dist = path.join(root, 'dist');
const output = path.join(dist, 'DAKE_STADIO-win32-x64');
const zip = path.join(dist, `DAKE_STADIO-${version}-win32-x64.zip`);
async function run() {
  await fs.access(path.join(output, 'DAKE_STADIO.exe'));
  for (const name of ['README.md', 'ORIGINAL.md', 'TEST_REPORT.md', 'release_body.md']) await fs.copyFile(path.join(root, name), path.join(output, name));
  const previous = path.join(dist, 'previous-builds', `DAKE_STADIO-${version}-before-docs-${new Date().toISOString().replace(/[:.]/g, '-')}.zip`);
  for (const candidate of [zip, previous]) {
    const relative = path.relative(root, path.resolve(candidate));
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) throw new Error('Unsafe distribution path');
  }
  await fs.mkdir(path.dirname(previous), { recursive: true });
  await fs.rename(zip, previous);
  const compression = spawnSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'Compress-Archive -LiteralPath $env:STADIO_PACKAGE_SOURCE -DestinationPath $env:STADIO_PACKAGE_ZIP -CompressionLevel Optimal'], {
    env: { ...process.env, STADIO_PACKAGE_SOURCE: output, STADIO_PACKAGE_ZIP: zip }, windowsHide: true, encoding: 'utf8'
  });
  if (compression.status !== 0) throw new Error(compression.stderr || 'ZIP rebuild failed');
  console.log(`Updated distribution docs and rebuilt ZIP: ${zip}`);
}
run().catch(error => { console.error(error); process.exitCode = 1; });
