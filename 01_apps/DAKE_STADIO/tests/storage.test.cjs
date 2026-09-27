'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const testTemporary = require('../scripts/local-test-env.cjs')().temp;
const path = require('node:path');
const storage = require('../desktop/storage.cjs');

const raster = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
const sample = (name = '日本語の作品') => ({
  format: 'dake-stadio', version: 1, width: 1200, height: 800, name, background: '#ffffff',
  canvas: { version: '7.4.0', objects: [{ type: 'Image', src: raster }, { type: 'Path', path: [['M', 0, 0], ['L', 30, 50]], name: '曲線' }] }
});

test('project round trip preserves embedded raster, vector path and Japanese metadata; backup retains previous successful version', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-storage-'));
  const destination = path.join(directory, '作品.dake');
  await storage.writeProject(destination, sample());
  assert.deepEqual(await storage.readProject(destination), sample());
  await storage.writeProject(destination, sample('変更後'));
  assert.deepEqual(await storage.readProject(destination), sample('変更後'));
  assert.deepEqual(await storage.readProject(`${destination}.bak`), sample());
  assert.equal((await fs.readdir(directory)).filter(name => name.endsWith('.tmp')).length, 0);
});

test('failed backup leaves destination bytes untouched and removes staging files', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-failure-'));
  const destination = path.join(directory, 'protected.dake');
  await storage.writeProject(destination, sample());
  await fs.mkdir(`${destination}.bak`);
  const original = await fs.readFile(destination);
  await assert.rejects(storage.writeProject(destination, sample('未保存')));
  assert.deepEqual(await fs.readFile(destination), original);
  assert.equal((await fs.readdir(directory)).filter(name => name.endsWith('.tmp')).length, 0);
});

test('invalid dimensions, unknown versions, excessive nesting and remote image URLs are rejected before disk mutation', async () => {
  assert.throws(() => storage.validateProject({ ...sample(), width: 9000 }), /DAKE STADIO/);
  assert.throws(() => storage.validateProject({ ...sample(), width: 8000, height: 8000 }), /DAKE STADIO/);
  assert.throws(() => storage.validateProject({ ...sample(), version: 999 }), /DAKE STADIO/);
  for (const resource of ['https://example.org/pixel.png', 'file:///C:/private.png', 'data:image/svg+xml;base64,AAAA']) {
    const project = sample();
    project.canvas.objects[0].src = resource;
    assert.throws(() => storage.validateProject(project), /外部参照/);
  }
  const malicious = JSON.parse('{"format":"dake-stadio","version":1,"width":1,"height":1,"canvas":{"objects":[]},"__proto__":{"polluted":true}}');
  assert.throws(() => storage.validateProject(malicious), /外部参照/);
  const nested = sample(); let target = nested;
  for (let i = 0; i < 90; i++) { target.child = {}; target = target.child; }
  assert.throws(() => storage.validateProject(nested), /DAKE STADIO/);
});

test('valid internal SVG references work and active or external SVG resources fail', () => {
  const good = '<svg xmlns="http://www.w3.org/2000/svg"><defs><linearGradient id="g" /></defs><path fill="url(#g)" d="M0 0 L9 9" /></svg>';
  assert.equal(storage.validateSvg(good), good);
  for (const body of ['<script>alert(1)</script>', '<foreignObject/>', '<image href="https://example.org/a.png"/>', '<path style="fill:url(file:///c:/secret)"/>', '<g onclick="alert(1)"/>', '<style>@import "https://example.org";</style>']) {
    assert.throws(() => storage.validateSvg(`<svg>${body}</svg>`), /SVG/);
  }
  assert.throws(() => storage.validateSvg('<!DOCTYPE svg [<!ENTITY bad SYSTEM "file:///secret">]><svg/>'), /SVG/);
});

test('export validates signatures and output mime; filename removes path traversal syntax', () => {
  assert.equal(storage.encodeExport('png', raster).subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
  assert.throws(() => storage.encodeExport('jpeg', raster), /書き出し/);
  assert.throws(() => storage.encodeExport('png', 'data:image/png;base64,SGVsbG8='), /書き出し/);
  assert.equal(storage.safeName('../危険/名前.png'), '.._危険_名前.png');
  assert.equal(storage.safeName('CON'), '_CON');
});

test('invalid project never replaces an existing valid file', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-invalid-'));
  const destination = path.join(directory, 'safe.dake');
  await storage.writeProject(destination, sample());
  await assert.rejects(storage.writeProject(destination, { ...sample(), version: 0 }));
  assert.deepEqual(await storage.readProject(destination), sample());
});

test('truncated JSON and oversized files are rejected', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-read-'));
  const truncated = path.join(directory, 'truncated.dake');
  await fs.writeFile(truncated, '{"format":"dake-stadio",');
  await assert.rejects(storage.readProject(truncated), /DAKE STADIO/);
  const large = path.join(directory, 'large.dake');
  const handle = await fs.open(large, 'w');
  await handle.truncate(storage.MAX_BYTES + 1);
  await handle.close();
  await assert.rejects(storage.readProject(large), /128 MiB/);
});

test('recovery falls back to previous valid state after truncated latest write; corrupt originals remain for diagnosis', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-recovery-'));
  const destination = path.join(directory, 'recovery.json');
  assert.equal(await storage.readRecovery(destination), null);
  await storage.writeRecovery(destination, sample('先の編集'));
  await storage.writeRecovery(destination, sample('最新の編集'));
  assert.equal((await storage.readRecovery(destination)).data.name, '最新の編集');
  await fs.writeFile(destination, '{"data":');
  assert.equal((await storage.readRecovery(destination)).data.name, '先の編集');
  assert.equal(await fs.readFile(destination, 'utf8'), '{"data":');
  await fs.writeFile(`${destination}.bak`, 'broken');
  await assert.rejects(storage.readRecovery(destination), /復旧データ/);
  assert.equal(await fs.readFile(`${destination}.bak`, 'utf8'), 'broken');
});

test('explicit recovery discard removes both recovery generations and is idempotent', async () => {
  const directory = await fs.mkdtemp(path.join(testTemporary, 'dake-stadio-discard-'));
  const destination = path.join(directory, 'recovery.json');
  await storage.writeRecovery(destination, sample());
  await storage.writeRecovery(destination, sample('次'));
  await storage.discardRecovery(destination);
  assert.equal(await storage.readRecovery(destination), null);
  await storage.discardRecovery(destination);
});
