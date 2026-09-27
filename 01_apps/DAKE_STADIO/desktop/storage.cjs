'use strict';

const fs = require('node:fs/promises');
const { constants } = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const UI_TEXT = require('../src/ui-text.json').desktop;
const MAX_BYTES = 128 * 1024 * 1024;
const RASTER_DATA = /^data:image\/(?:png|jpeg|webp|gif|bmp);base64,[A-Za-z0-9+/\r\n]*={0,2}$/;
const RESOURCE_KEYS = new Set(['src', 'source', 'href', 'xlink:href', 'url']);

function fail(key, code) {
  const error = new Error(UI_TEXT[key]);
  error.code = code;
  throw error;
}

function validateExtras(data) {
  const meta = data.meta || {}, dpi = meta.dpi ?? 96, bleed = meta.bleed ?? 0, safe = meta.safe ?? 0;
  if (!Number.isFinite(dpi) || dpi < 36 || dpi > 2400 || !['px', 'mm'].includes(meta.unit || 'px') || !Number.isFinite(bleed) || !Number.isFinite(safe) || bleed < 0 || safe < 0 || 2 * (bleed + safe) >= Math.min(data.width, data.height)) fail('errorProject', 'INVALID_LAYOUT');
  for (const axis of ['horizontal', 'vertical']) {
    const guides = meta.guides?.[axis] || [];
    if (!Array.isArray(guides) || guides.length > 200 || guides.some(value => !Number.isFinite(value) || value < 0 || value > (axis === 'horizontal' ? data.height : data.width))) fail('errorProject', 'INVALID_LAYOUT');
  }
  if (meta.snap?.gridSize !== undefined && (!Number.isFinite(meta.snap.gridSize) || meta.snap.gridSize < 1 || meta.snap.gridSize > 4096)) fail('errorProject', 'INVALID_LAYOUT');
  const fonts = data.fonts || [];
  if (!Array.isArray(fonts) || fonts.length > 32) fail('errorProject', 'INVALID_FONTS');
  let embeddedSize = 0;
  for (const font of fonts) {
    if (!font || typeof font.family !== 'string' || font.family.length > 200 || !['google', 'local'].includes(font.source)) fail('errorProject', 'INVALID_FONTS');
    if (font.source === 'local' && font.data !== undefined) fail('errorFontLicense', 'LOCAL_FONT_EMBED');
    if (font.source === 'google') {
      if (typeof font.data !== 'string' || !/^data:(?:font\/(?:ttf|otf|sfnt)|application\/(?:font-sfnt|octet-stream));base64,[A-Za-z0-9+/]*={0,2}$/.test(font.data) || font.data.length > 48 * 1024 * 1024) fail('errorProject', 'INVALID_FONTS');
      embeddedSize += font.data.length;
      if (typeof font.license !== 'string' || font.license.length > 128000 || !/SIL OPEN FONT LICENSE|Apache License|Ubuntu Font Licen[cs]e/i.test(font.license)) fail('errorFontLicense', 'INVALID_FONT_LICENSE');
      const bytes = Buffer.from(font.data.slice(font.data.indexOf(',') + 1), 'base64');
      if (!/^[a-f0-9]{64}$/i.test(font.sha256 || '') || crypto.createHash('sha256').update(bytes).digest('hex') !== font.sha256) fail('errorFontInvalid', 'FONT_HASH_MISMATCH');
      if (bytes.length < 12 || !(bytes.readUInt32BE(0) === 0x00010000 || bytes.toString('ascii', 0, 4) === 'OTTO' || bytes.toString('ascii', 0, 4) === 'true')) fail('errorFontInvalid', 'INVALID_FONT_FORMAT');
    }
  }
  if (embeddedSize > 96 * 1024 * 1024) fail('errorLarge', 'TOO_LARGE');
}

function validateProject(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data) || data.format !== 'dake-stadio' || ![1, 2].includes(data.version) || !data.canvas || !Array.isArray(data.canvas.objects)) fail('errorProject', 'INVALID_PROJECT');
  if (!Number.isInteger(data.width) || !Number.isInteger(data.height) || data.width < 1 || data.height < 1 || data.width > 8192 || data.height > 8192 || data.width * data.height > 32000000) fail('errorProject', 'INVALID_SIZE');
  if (data.canvas.objects.length > 2000) fail('errorProject', 'TOO_MANY_OBJECTS');
  validateExtras(data);
  let entries = 0;
  function visit(value, depth, inFonts = false) {
    if (depth > 80 || ++entries > 500000) fail('errorProject', 'COMPLEX_PROJECT');
    if (typeof value === 'number' && !Number.isFinite(value)) fail('errorProject', 'INVALID_NUMBER');
    if (!value || typeof value !== 'object') return;
    for (const [key, child] of Object.entries(value)) {
      if (key === '__proto__' || key === 'constructor' || key === 'prototype') fail('errorUnsafe', 'UNSAFE_PROJECT');
      if (!(inFonts && key === 'source' && ['local', 'google'].includes(child)) && RESOURCE_KEYS.has(key.toLowerCase()) && typeof child === 'string' && child && !RASTER_DATA.test(child) && !/^#[\w:.-]+$/.test(child)) fail('errorUnsafe', 'EXTERNAL_RESOURCE');
      visit(child, depth + 1, inFonts || (value === data && key === 'fonts'));
    }
  }
  visit(data, 0);
  const encoded = JSON.stringify(data);
  if (Buffer.byteLength(encoded) > MAX_BYTES) fail('errorLarge', 'TOO_LARGE');
  return data;
}

function validateSvg(svg) {
  if (typeof svg !== 'string' || Buffer.byteLength(svg) > MAX_BYTES || !/<svg(?:\s|>)/i.test(svg)) fail('errorSvg', 'INVALID_SVG');
  if (/<\s*(?:script|foreignObject|iframe|object|embed|audio|video|image-set)\b/i.test(svg) || /<!\s*(?:DOCTYPE|ENTITY)/i.test(svg) || /\bon[a-z]+\s*=/i.test(svg) || /@import\b/i.test(svg) || /(?:javascript|vbscript)\s*:/i.test(svg)) fail('errorSvg', 'UNSAFE_SVG');
  for (const match of svg.matchAll(/\b(?:href|xlink:href)\s*=\s*(["'])(.*?)\1/gis)) {
    if (!/^#[\w:.-]+$/.test(match[2]) && !RASTER_DATA.test(match[2])) fail('errorSvg', 'EXTERNAL_SVG_RESOURCE');
  }
  for (const match of svg.matchAll(/url\s*\(\s*(["']?)(.*?)\1\s*\)/gis)) {
    if (!/^#[\w:.-]+$/.test(match[2].trim())) fail('errorSvg', 'EXTERNAL_SVG_RESOURCE');
  }
  // XML entity references in URLs could hide a remote protocol. Only plain
  // fragment identifiers and base64 data are accepted by the checks above.
  return svg;
}

async function readLimited(filePath) {
  const info = await fs.stat(filePath);
  if (!info.isFile() || info.size > MAX_BYTES) fail('errorLarge', 'TOO_LARGE');
  const buffer = await fs.readFile(filePath);
  if (buffer.length > MAX_BYTES) fail('errorLarge', 'TOO_LARGE');
  return buffer;
}

async function readProject(filePath) {
  const bytes = await readLimited(filePath);
  let data;
  try { data = JSON.parse(bytes.toString('utf8').replace(/^\uFEFF/, '')); }
  catch { fail('errorProject', 'INVALID_JSON'); }
  return validateProject(data);
}

async function writeSynced(tempPath, bytes) {
  const handle = await fs.open(tempPath, 'wx', 0o600);
  try { await handle.writeFile(bytes); await handle.sync(); }
  finally { await handle.close(); }
}

async function atomicWrite(filePath, bytes, { backup = false } = {}) {
  const parent = path.dirname(filePath);
  const suffix = `.stadio-${process.pid}-${crypto.randomBytes(8).toString('hex')}.tmp`;
  const temp = path.join(parent, `${path.basename(filePath)}${suffix}`);
  const backupTemp = path.join(parent, `${path.basename(filePath)}.bak${suffix}`);
  try {
    await writeSynced(temp, bytes);
    if (backup) {
      let copied = false;
      try { await fs.copyFile(filePath, backupTemp, constants.COPYFILE_EXCL); copied = true; }
      catch (error) { if (error.code !== 'ENOENT') throw error; }
      if (copied) {
        const backupHandle = await fs.open(backupTemp, 'r+');
        try { await backupHandle.sync(); } finally { await backupHandle.close(); }
        await fs.rename(backupTemp, `${filePath}.bak`);
      }
    }
    await fs.rename(temp, filePath);
    // Windows does not allow fsync on directory handles. The file content is
    // synced before the atomic rename; directory sync is best effort elsewhere.
    if (process.platform !== 'win32') {
      const directory = await fs.open(parent, 'r');
      try { await directory.sync(); } finally { await directory.close(); }
    }
  } finally {
    await fs.unlink(temp).catch(() => {});
    await fs.unlink(backupTemp).catch(() => {});
  }
}

async function writeProject(filePath, data) {
  validateProject(data);
  await atomicWrite(filePath, JSON.stringify(data), { backup: true });
}

async function writeRecovery(filePath, data) {
  validateProject(data);
  await atomicWrite(filePath, JSON.stringify({ timestamp: new Date().toISOString(), data }), { backup: true });
}

async function readRecovery(filePath) {
  let found = false;
  for (const candidate of [filePath, `${filePath}.bak`]) {
    try {
      const info = await fs.stat(candidate);
      found = true;
      if (!info.isFile() || info.size > MAX_BYTES + 1024) continue;
      const bytes = await fs.readFile(candidate);
      if (bytes.length > MAX_BYTES + 1024) continue;
      const recovery = JSON.parse(bytes.toString('utf8'));
      validateProject(recovery.data);
      return { data: recovery.data, timestamp: recovery.timestamp || info.mtime.toISOString() };
    } catch (error) {
      if (error.code === 'ENOENT') continue;
      if (error.code && /^E[A-Z]+$/.test(error.code)) throw error;
    }
  }
  if (found) fail('errorDamagedRecovery', 'DAMAGED_RECOVERY');
  return null;
}

async function discardRecovery(filePath) {
  for (const candidate of [filePath, `${filePath}.bak`]) {
    await fs.unlink(candidate).catch(error => { if (error.code !== 'ENOENT') throw error; });
  }
}

function encodeExport(format, data) {
  if (format === 'svg') return Buffer.from(validateSvg(data), 'utf8');
  const mime = { png: 'png', jpeg: 'jpeg', jpg: 'jpeg', webp: 'webp' }[format];
  if (!mime || typeof data !== 'string' || !data.startsWith(`data:image/${mime};base64,`)) fail('errorExport', 'INVALID_EXPORT');
  const payload = data.slice(data.indexOf(',') + 1);
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(payload) || payload.length % 4 !== 0) fail('errorExport', 'INVALID_EXPORT');
  const bytes = Buffer.from(payload, 'base64');
  if (bytes.length > MAX_BYTES) fail('errorLarge', 'TOO_LARGE');
  const valid = (mime === 'png' && bytes.subarray(0, 8).equals(Buffer.from([137,80,78,71,13,10,26,10]))) ||
    (mime === 'jpeg' && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255) ||
    (mime === 'webp' && bytes.subarray(0, 4).toString() === 'RIFF' && bytes.subarray(8, 12).toString() === 'WEBP');
  if (!valid) fail('errorExport', 'INVALID_EXPORT');
  return bytes;
}

function safeName(value) {
  const name = String(value || UI_TEXT.defaultName).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/[. ]+$/, '').slice(0, 100) || UI_TEXT.defaultName;
  return /^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(name) ? `_${name}` : name;
}

module.exports = { MAX_BYTES, RASTER_DATA, validateProject, validateSvg, readLimited, readProject, atomicWrite, writeProject, readRecovery, writeRecovery, discardRecovery, encodeExport, safeName };
