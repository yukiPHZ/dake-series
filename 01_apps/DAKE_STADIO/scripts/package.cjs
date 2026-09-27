'use strict';
const path = require('node:path');
const fs = require('node:fs/promises');
const { spawnSync } = require('node:child_process');

async function packageApp() {
  const { root, temp, electronCache } = require('./local-test-env.cjs')();
  const metadata = require(path.join(root, 'package.json'));
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const stage = path.join(root, 'artifacts', `package-input-${stamp}`);
  const dist = path.join(root, 'dist');
  const expectedOutput = path.join(dist, 'DAKE_STADIO-win32-x64');
  const assertInside = candidate => {
    const relative = path.relative(root, path.resolve(candidate));
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) throw new Error('Package path is outside the project');
  };
  assertInside(stage);
  assertInside(expectedOutput);
  await fs.mkdir(stage, { recursive: true });
  await fs.mkdir(dist, { recursive: true });
  for (const directory of ['build', 'desktop', 'assets']) {
    await fs.cp(path.join(root, directory), path.join(stage, directory), { recursive: true });
  }
  await fs.mkdir(path.join(stage, 'src'), { recursive: true });
  await fs.copyFile(path.join(root, 'src', 'ui-text.json'), path.join(stage, 'src', 'ui-text.json'));
  await fs.writeFile(path.join(stage, 'package.json'), JSON.stringify({
    name: metadata.name, productName: 'DAKE STADIO', version: metadata.version,
    description: metadata.description, author: metadata.author,
    main: 'desktop/main.cjs', private: true
  }, null, 2));
  try {
    await fs.access(expectedOutput);
    const archive = path.join(dist, 'previous-builds', `DAKE_STADIO-win32-x64-${stamp}`);
    assertInside(archive);
    await fs.mkdir(path.dirname(archive), { recursive: true });
    await fs.rename(expectedOutput, archive);
  } catch (error) { if (error.code !== 'ENOENT') throw error; }

  const { packager } = await import('@electron/packager');
  const outputs = await packager({
    tmpdir: path.join(temp, 'packager'), download: { cacheRoot: electronCache },
    dir: stage, out: dist, name: 'DAKE_STADIO', executableName: 'DAKE_STADIO',
    platform: 'win32', arch: 'x64', electronVersion: metadata.devDependencies.electron,
    appVersion: metadata.version, buildVersion: metadata.version,
    appCopyright: 'Copyright (c) 2026 Yukihiko Kikuta',
    icon: path.join(root, 'assets', 'dake_icon.ico'),
    asar: true, prune: false, overwrite: false,
    win32metadata: { CompanyName: 'DAKE', FileDescription: 'DAKE STADIO', ProductName: 'DAKE STADIO', InternalName: 'DAKE_STADIO' }
  });
  for (const output of outputs) {
    for (const name of await fs.readdir(root)) {
      if (/\.md$/i.test(name) || /^LICENSE(?:\.txt)?$/i.test(name)) {
        await fs.copyFile(path.join(root, name), path.join(output, name));
      }
    }
    await fs.copyFile(path.join(root, 'assets', 'THIRD_PARTY_NOTICES.txt'), path.join(output, 'THIRD_PARTY_NOTICES.txt'));
    await fs.mkdir(path.join(output, 'examples'), { recursive: true });
    await fs.copyFile(path.join(root, 'evidence', 'sample-artwork.dake'), path.join(output, 'examples', 'はじめの作品.dake'));
    if (metadata.version !== '0.1.0') {
      const results = JSON.parse(await fs.readFile(path.join(root, 'evidence', 'production-results.json'), 'utf8'));
      if (!results.passed || typeof results.output !== 'string') throw new Error('Production examples have not passed verification');
      const exampleSource = path.resolve(results.output);
      if (!exampleSource.startsWith(path.join(root, 'test-output') + path.sep)) throw new Error('Example source must be inside canonical test-output');
      for (const kind of ['banner', 'flyer', 'logo', 'card']) {
        for (const extension of ['dake', 'png', 'svg', 'pdf']) {
          const name = kind + '.' + extension;
          const source = path.join(exampleSource, extension === 'dake' ? kind + '-edited.dake' : name);
          if (!(await fs.stat(source)).isFile()) throw new Error('Missing validated example: ' + name);
          await fs.copyFile(source, path.join(output, 'examples', name));
        }
      }
    }
    const zip = path.join(dist, `DAKE_STADIO-${metadata.version}-win32-x64.zip`);
    try {
      await fs.access(zip);
      const previousZip = path.join(dist, 'previous-builds', `DAKE_STADIO-${metadata.version}-win32-x64-${stamp}.zip`);
      assertInside(previousZip);
      await fs.mkdir(path.dirname(previousZip), { recursive: true });
      await fs.rename(zip, previousZip);
    } catch (error) { if (error.code !== 'ENOENT') throw error; }
    const compressed = spawnSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'Compress-Archive -LiteralPath $env:STADIO_PACKAGE_SOURCE -DestinationPath $env:STADIO_PACKAGE_ZIP -CompressionLevel Optimal'], {
      env: { ...process.env, STADIO_PACKAGE_SOURCE: output, STADIO_PACKAGE_ZIP: zip },
      encoding: 'utf8', windowsHide: true
    });
    if (compressed.status !== 0) throw new Error(compressed.stderr || 'ZIP creation failed');
    console.log(`Executable: ${path.join(output, 'DAKE_STADIO.exe')}`);
    console.log(`Portable ZIP: ${zip}`);
  }
}

packageApp().catch(error => { console.error(error); process.exitCode = 1; });
