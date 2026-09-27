require('./local-test-env.cjs')();
'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const assert = require('node:assert/strict');
const esbuild = require('esbuild');
const root = path.resolve(__dirname, '..');
const digest = bytes => createHash('sha256').update(bytes).digest('hex');
async function run() {
  const archive = path.join(root, 'dist', 'DAKE_STADIO-win32-x64', 'resources', 'app.asar');
  const asar = await import('@electron/asar');
  const rebuilt = await esbuild.build({
    absWorkingDir: root, entryPoints: ['src/app.js'], bundle: true,
    outfile: path.join(root, 'build', 'app.js'), format: 'iife', platform: 'browser',
    target: 'chrome140', tsconfigRaw: {}, minify: false, sourcemap: false,
    legalComments: 'eof', logLevel: 'silent', write: false
  });
  assert.equal(digest(rebuilt.outputFiles[0].contents), digest(await fs.readFile(path.join(root, 'build', 'app.js'))), 'Current app/engine source must reproduce shipped build byte for byte');
  const vector = await esbuild.build({ absWorkingDir: root, entryPoints: ['src/vector-worker.js'], bundle: true, outfile: path.join(root, 'build', 'vector-worker.js'), format: 'iife', platform: 'browser', target: 'chrome140', tsconfigRaw: {}, minify: false, sourcemap: false, legalComments: 'eof', logLevel: 'silent', write: false });
  assert.equal(digest(vector.outputFiles[0].contents), digest(await fs.readFile(path.join(root, 'build', 'vector-worker.js'))), 'Current vector worker source must reproduce shipped build byte for byte');
  const files = {};
  for (const directory of ['build', 'desktop', 'assets']) {
    for (const entry of await fs.readdir(path.join(root, directory), { withFileTypes: true })) {
      if (!entry.isFile()) throw new Error(`Unexpected nested runtime directory: ${directory}/${entry.name}`);
      const relative = `${directory}/${entry.name}`;
      const expected = await fs.readFile(path.join(root, directory, entry.name));
      const actual = asar.extractFile(archive, relative);
      assert.equal(digest(actual), digest(expected), `Bundled runtime mismatch: ${relative}`);
      files[relative] = digest(actual);
    }
  }
  const texts = await fs.readFile(path.join(root, 'src', 'ui-text.json'));
  assert.equal(digest(asar.extractFile(archive, 'src/ui-text.json')), digest(texts));
  files['src/ui-text.json'] = digest(texts);
  for (const name of ['index.html', 'style.css', 'raster-worker.js']) assert.equal(digest(await fs.readFile(path.join(root, 'src', name))), digest(await fs.readFile(path.join(root, 'build', name))), `Uncopied source change: ${name}`);
  const expectedEntries = [...Object.keys(files), 'package.json'].sort();
  const actualEntries = asar.listPackage(archive).map(name => name.replace(/^[/\\]/, '').replace(/\\/g, '/')).filter(name => !['assets', 'build', 'desktop', 'src'].includes(name)).sort();
  assert.deepEqual(actualEntries, expectedEntries, 'Package must contain only audited runtime files');
  const report = { passed: true, checks: ['final app, engine and vector worker sources reproduce build byte-for-byte', 'every bundled build desktop asset and UI_TEXT file matches workspace', 'static build resources match current source', 'app.asar contains only audited runtime files'], archive, appAsarSha256: digest(await fs.readFile(archive)), files };
  await fs.writeFile(path.join(root, 'evidence', 'source-package-consistency.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}
run().catch(error => { console.error(error); process.exitCode = 1; });
