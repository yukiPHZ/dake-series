require('./local-test-env.cjs')();
'use strict';
const fs = require('node:fs/promises');
const { createReadStream } = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { spawnSync } = require('node:child_process');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const version = require('../package.json').version;
const source = path.join(root, 'dist', 'DAKE_STADIO-win32-x64');
const zip = path.join(root, 'dist', `DAKE_STADIO-${version}-win32-x64.zip`);
const output = path.join(root, 'test-output', `zip-extracted-${Date.now()}`);
const extracted = path.join(output, 'DAKE_STADIO-win32-x64');
async function hash(filePath) {
  const digest = createHash('sha256');
  for await (const chunk of createReadStream(filePath)) digest.update(chunk);
  return digest.digest('hex');
}
async function files(directory, prefix = '') {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const result = [];
  for (const entry of entries) {
    const relative = path.join(prefix, entry.name);
    if (entry.isDirectory()) result.push(...await files(path.join(directory, entry.name), relative));
    else if (entry.isFile()) result.push(relative);
    else throw new Error(`Unexpected non-regular distribution entry: ${relative}`);
  }
  return result.sort();
}
async function run() {
  await fs.mkdir(output, { recursive: true });
  const script = "Add-Type -AssemblyName System.IO.Compression.FileSystem; $stadioArchive = [System.IO.Compression.ZipFile]::OpenRead($env:STADIO_VERIFY_ZIP); try { foreach ($stadioEntry in $stadioArchive.Entries) { $stadioEntryPath = $stadioEntry.FullName.Replace([char]92, [char]47); if ([System.IO.Path]::IsPathRooted($stadioEntryPath) -or ($stadioEntryPath.Split([char]47) -contains '..')) { throw 'Unsafe archive path' } } } finally { $stadioArchive.Dispose() }; Expand-Archive -LiteralPath $env:STADIO_VERIFY_ZIP -DestinationPath $env:STADIO_VERIFY_OUTPUT";
  const expansion = spawnSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script], {
    env: { ...process.env, STADIO_VERIFY_ZIP: zip, STADIO_VERIFY_OUTPUT: output }, windowsHide: true, encoding: 'utf8'
  });
  if (expansion.status !== 0) throw new Error(expansion.stderr || 'Archive extraction failed');
  const originalFiles = await files(source);
  const extractedFiles = await files(extracted);
  assert.deepEqual(extractedFiles, originalFiles);
  const fingerprints = {};
  for (const relative of originalFiles) {
    const originalHash = await hash(path.join(source, relative));
    const extractedHash = await hash(path.join(extracted, relative));
    assert.equal(extractedHash, originalHash, `ZIP content mismatch: ${relative}`);
    fingerprints[relative] = originalHash;
  }
  const report = { passed: true, zip, extracted, fileCount: originalFiles.length, zipBytes: (await fs.stat(zip)).size, zipSha256: await hash(zip), checks: ['archive entries contain no path traversal', 'all extracted file names match distribution', 'every extracted file SHA256 matches original distribution'], fingerprints };
  await fs.writeFile(path.join(root, 'evidence', 'archive-results.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify({ ...report, fingerprints: undefined }, null, 2));
}
run().catch(error => { console.error(error); process.exitCode = 1; });
