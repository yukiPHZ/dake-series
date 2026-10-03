'use strict';
const fs = require('node:fs/promises'), path = require('node:path'), https = require('node:https'), crypto = require('node:crypto');
const { atomicWrite } = require('./storage.cjs');
const { extractFont, MAX_FONT_BYTES } = require('./font-binary.cjs');
const UI_TEXT = require('../src/ui-text.json').desktop;
const hash = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const fail = key => new Error(UI_TEXT[key] || key);
function createFontService({ userData, catalogPath, progress = () => {}, offline = process.env.STADIO_FONT_OFFLINE === '1' }) {
  const cache = path.join(userData, 'fonts'), jobs = new Map();
  let catalogPromise;
  const readCatalog = () => catalogPromise ||= fs.readFile(catalogPath, 'utf8').then(JSON.parse);
  async function cachedMetadata() {
    const directory = path.join(cache, 'descriptors'), result = [];
    for (const name of await fs.readdir(directory).catch(() => [])) {
      if (!/^[a-f0-9]{64}\.json$/.test(name)) continue;
      try { const value = JSON.parse(await fs.readFile(path.join(directory, name), 'utf8')); if (value.source === 'google' && typeof value.family === 'string') result.push({...value, cacheKey:name.slice(0,-5)}); } catch {}
    }
    return result;
  }
  async function list() {
    const catalog = await readCatalog(), cached = new Set((await cachedMetadata()).map(value => value.family));
    return { ...catalog, families: catalog.families.map(family => ({ ...family, id: family.id || family.directory, cached: cached.has(family.family) })) };
  }
  function permittedUrl(commit, file) {
    if (!/^[a-f0-9]{40}$/.test(commit) || typeof file !== 'string' || !/^(?:ofl|apache|ufl)\/[a-z0-9_-]+\//i.test(file) || file.split('/').some(part => part === '..' || !part)) throw fail('errorFontDownload');
    return 'https://raw.githubusercontent.com/google/fonts/' + commit + '/' + file.split('/').map(encodeURIComponent).join('/');
  }
  function fetchBytes(url, maximum, signal, requestId, phase) {
    return new Promise((resolve, reject) => {
      const request = https.get(url, { signal, headers: { 'User-Agent': 'DAKE-STADIO/0.2.0' } }, response => {
        if (response.statusCode !== 200) { response.resume(); reject(fail('errorFontDownload')); return; }
        const declared = Number(response.headers['content-length'] || 0);
        if (declared > maximum) { response.destroy(); reject(fail('errorFontDownload')); return; }
        let loaded = 0, lastNotice = 0; const chunks = [];
        response.on('data', chunk => {
          loaded += chunk.length;
          if (loaded > maximum) { response.destroy(fail('errorFontDownload')); return; }
          chunks.push(chunk);
          if (Date.now() - lastNotice > 100) { progress({ id: requestId, phase, loaded, total: declared || null }); lastNotice = Date.now(); }
        });
        response.once('end', () => { progress({ id: requestId, phase, loaded, total: declared || loaded }); resolve(Buffer.concat(chunks)); });
        response.once('error', reject);
      });
      request.setTimeout(30000, () => request.destroy(fail('errorFontDownload')));
      request.once('error', reject);
    });
  }
  async function readCached(key) {
    try {
      const value = JSON.parse(await fs.readFile(path.join(cache, 'descriptors', key + '.json'), 'utf8'));
      if (!/^[a-f0-9]{64}$/.test(value.sha256) || !/^[a-f0-9]{64}$/.test(value.licenseSha256)) return null;
      const bytes = await fs.readFile(path.join(cache, 'blobs', value.sha256 + '.sfnt'));
      const license = await fs.readFile(path.join(cache, 'licenses', value.licenseSha256 + '.txt'), 'utf8');
      if (hash(bytes) !== value.sha256 || hash(Buffer.from(license)) !== value.licenseSha256) return null;
      const inspected = extractFont(bytes, value.postscriptName);
      return { ...value, coverage: inspected.coverage, license, data: 'data:font/' + value.format + ';base64,' + bytes.toString('base64'), cached: true };
    } catch { return null; }
  }
  async function download(request) {
    const catalog = await readCatalog();
    const family = catalog.families.find(value => value.family === request?.family || (value.id || value.directory) === request?.id);
    if (!family || !Array.isArray(family.files) || !family.files.length) throw fail('errorFontUnavailable');
    const style = request.style === 'italic' ? 'italic' : 'normal';
    const weight = Number.isFinite(Number(request.weight)) ? Math.max(1, Math.min(1000, Number(request.weight))) : 400;
    const compatible = family.files.filter(file => (file.style === 'italic' ? 'italic' : 'normal') === style);
    const candidates = compatible.length ? compatible : family.files;
    const file = [...candidates].sort((a,b) => Math.abs((Number(a.weight) || 400) - weight) - Math.abs((Number(b.weight) || 400) - weight))[0];
    if (!/\.(?:ttf|otf)$/i.test(file.path)) throw fail('errorFontUnavailable');
    if (Number(file.size) > 0 && Math.ceil(Number(file.size) / 3) * 4 + 32 > 48 * 1024 * 1024) throw fail('errorFontCapacity');
    const fontUrl = permittedUrl(catalog.sourceCommit, file.path), licenseUrl = permittedUrl(catalog.sourceCommit, family.licensePath);
    const key = hash(Buffer.from([catalog.sourceCommit, file.path, weight, style].join('|')));
    const cached = await readCached(key); if (cached) return cached;
    if (offline) throw fail('errorFontDownload');
    const requestId = String(request.requestId || request.downloadId || request.id || family.directory).slice(0, 200);
    if (jobs.has(requestId)) throw fail('errorFontDownload');
    const controller = new AbortController(); jobs.set(requestId, controller);
    try {
      const [fontBytes, licenseBytes] = await Promise.all([
        fetchBytes(fontUrl, MAX_FONT_BYTES, controller.signal, requestId, 'font'),
        fetchBytes(licenseUrl, 256 * 1024, controller.signal, requestId, 'license')
      ]);
      const license = licenseBytes.toString('utf8');
      if (!/SIL OPEN FONT LICENSE|Apache License|Ubuntu Font Licen[cs]e/i.test(license)) throw fail('errorFontLicense');
      const gitBlob = bytes => crypto.createHash('sha1').update(Buffer.from('blob ' + bytes.length + '\0')).update(bytes).digest('hex');
      if ((file.blobSha && gitBlob(fontBytes) !== file.blobSha) || (family.licenseBlobSha && gitBlob(licenseBytes) !== family.licenseBlobSha)) throw fail('errorFontDownload');
      const result = extractFont(fontBytes, file.postscriptName);
      if (Math.ceil(result.bytes.length / 3) * 4 + 32 > 48 * 1024 * 1024) throw fail('errorFontCapacity');
      if (controller.signal.aborted) throw fail('errorFontCancelled');
      const actualWeight = result.axes.wght ? Math.max(result.axes.wght.min, Math.min(result.axes.wght.max, weight)) : result.defaultWeight;
      const actualStyle = result.style;
      const descriptor = {
        id: 'google:' + family.directory + ':' + actualStyle + ':' + actualWeight,
        family: family.family, style: actualStyle, weight: actualWeight, weightRange: result.axes.wght ? result.axes.wght.min + ' ' + result.axes.wght.max : String(actualWeight), source: 'google', format: result.format,
        sha256: result.sha256, sourceSha256: hash(fontBytes), licenseSha256: hash(licenseBytes),
        licenseType: family.licenseType, version: catalog.sourceCommit, fontUrl, licenseUrl,
        postscriptName: result.postscriptNames[0] || '', fsType: result.fsType, coverage: result.coverage
      };
      for (const directory of ['blobs', 'licenses', 'descriptors']) await fs.mkdir(path.join(cache, directory), { recursive: true });
      await atomicWrite(path.join(cache, 'blobs', descriptor.sha256 + '.sfnt'), result.bytes);
      await atomicWrite(path.join(cache, 'licenses', descriptor.licenseSha256 + '.txt'), licenseBytes);
      if (controller.signal.aborted) throw fail('errorFontCancelled');
      // Publish descriptor last; incomplete transfers never appear as cached fonts.
      await atomicWrite(path.join(cache, 'descriptors', key + '.json'), JSON.stringify(descriptor));
      return { ...descriptor, license, data: 'data:font/' + descriptor.format + ';base64,' + result.bytes.toString('base64'), cached: false };
    } catch (error) {
      const cancelled = controller.signal.aborted;
      controller.abort();
      if (cancelled || error.name === 'AbortError') throw fail('errorFontCancelled');
      if (error.message === 'FONT_INVALID') throw fail('errorFontInvalid');
      if (error.message === UI_TEXT.errorFontLicense || error.message === UI_TEXT.errorFontCapacity) throw error;
      throw fail('errorFontDownload');
    } finally { jobs.delete(requestId); }
  }
  function cancel(id) { const job = jobs.get(String(id)); if (job) job.abort(); return !!job; }
  function prepareLocal({ descriptor, bytes }) {
    if (!descriptor || descriptor.source !== 'local' || typeof descriptor.postscriptName !== 'string' || descriptor.postscriptName.length > 250) throw fail('errorFontInvalid');
    let result;
    try { result = extractFont(Buffer.from(bytes), descriptor.postscriptName); } catch { throw fail('errorFontInvalid'); }
    return { ...descriptor, weight: result.defaultWeight, style: result.style, weightRange: result.axes.wght ? result.axes.wght.min + ' ' + result.axes.wght.max : String(result.defaultWeight), coverage: result.coverage, fsType: result.fsType, sha256: result.sha256, format: result.format, data: 'data:font/' + result.format + ';base64,' + result.bytes.toString('base64') };
  }
  async function listCached() {
    const result=[];
    for (const descriptor of await cachedMetadata()) {
      const valid=await readCached(descriptor.cacheKey);
      if(valid){const {data,license,...metadata}=valid;result.push({...metadata,cacheKey:descriptor.cacheKey});}
    }
    return result;
  }
  async function getCached(key) {
    if(typeof key!=='string'||!/^[a-f0-9]{64}$/.test(key))throw fail('errorFontUnavailable');
    const font=await readCached(key);if(!font)throw fail('errorFontUnavailable');return font;
  }
  return { list, download, cancel, prepareLocal, cachedMetadata, listCached, getCached };
}
module.exports = { createFontService };

