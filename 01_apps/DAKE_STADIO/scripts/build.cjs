'use strict';
const path = require('node:path');
const fs = require('node:fs/promises');
const esbuild = require('esbuild');

async function build() {
  const root = path.resolve(__dirname, '..');
  const output = path.join(root, 'build');
  await fs.mkdir(output, { recursive: true });
  await esbuild.build({
    absWorkingDir: root,
    entryPoints: ['src/app.js'],
    bundle: true, outfile: path.join(output, 'app.js'),
    format: 'iife', platform: 'browser', target: 'chrome140',
    tsconfigRaw: {},
    minify: false, sourcemap: false, legalComments: 'eof',
    logLevel: 'info'
  });
  await esbuild.build({ absWorkingDir: root, entryPoints: ['src/vector-worker.js'], bundle: true, outfile: path.join(output, 'vector-worker.js'), format: 'iife', platform: 'browser', target: 'chrome140', tsconfigRaw: {}, minify: false, sourcemap: false, legalComments: 'eof', logLevel: 'info' });
  for (const name of ['index.html', 'style.css', 'raster-worker.js']) await fs.copyFile(path.join(root, 'src', name), path.join(output, name));
}

if (require.main === module) build().catch(error => { console.error(error); process.exitCode = 1; });
module.exports = build;
