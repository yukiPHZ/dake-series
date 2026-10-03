'use strict';
const crypto = require('node:crypto');
const MAX_FONT_BYTES = 64 * 1024 * 1024;
function invalid() { const error = new Error('FONT_INVALID'); error.code = 'FONT_INVALID'; throw error; }
function checkRange(bytes, offset, length) { if (!Number.isInteger(offset) || !Number.isInteger(length) || offset < 0 || length < 0 || offset + length > bytes.length) invalid(); }
function tables(bytes, faceOffset) {
  checkRange(bytes, faceOffset, 12);
  const magic = bytes.toString('ascii', faceOffset, faceOffset + 4);
  if (bytes.readUInt32BE(faceOffset) !== 0x00010000 && !['OTTO', 'true'].includes(magic)) invalid();
  const count = bytes.readUInt16BE(faceOffset + 4);
  if (!count || count > 200) invalid();
  checkRange(bytes, faceOffset + 12, count * 16);
  const result = new Map();
  for (let i = 0; i < count; i++) {
    const at = faceOffset + 12 + i * 16, tag = bytes.toString('ascii', at, at + 4);
    const offset = bytes.readUInt32BE(at + 8), length = bytes.readUInt32BE(at + 12);
    checkRange(bytes, offset, length);
    if (result.has(tag)) invalid();
    result.set(tag, { tag, offset, length });
  }
  return { magic: bytes.readUInt32BE(faceOffset), records: result };
}
function nameValues(bytes, record, wantedId) {
  if (!record || record.length < 6) return [];
  const start = record.offset, count = bytes.readUInt16BE(start + 2), strings = start + bytes.readUInt16BE(start + 4);
  if (6 + count * 12 > record.length) invalid();
  const result = [];
  for (let i = 0; i < count; i++) {
    const at = start + 6 + i * 12;
    if (bytes.readUInt16BE(at + 6) !== wantedId) continue;
    const platform = bytes.readUInt16BE(at), length = bytes.readUInt16BE(at + 8), offset = strings + bytes.readUInt16BE(at + 10);
    checkRange(bytes, offset, length);
    if (offset + length > start + record.length) invalid();
    const value = Buffer.from(bytes.subarray(offset, offset + length));
    if (platform === 0 || platform === 3) {
      if (length % 2) invalid(); value.swap16(); result.push(value.toString('utf16le'));
    } else if (platform === 1) result.push(value.toString('latin1'));
  }
  return result;
}
function checksum(bytes) {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) {
    let word = 0;
    for (let n = 0; n < 4; n++) word = ((word << 8) | (bytes[i + n] || 0)) >>> 0;
    sum = (sum + word) >>> 0;
  }
  return sum;
}

function unicodeCoverage(bytes, record) {
  if (!record || record.length < 4) return [];
  const start = record.offset, count = bytes.readUInt16BE(start + 2), points = new Set(); let examined = 0;
  if (count > 256 || 4 + count * 8 > record.length) invalid();
  for (let i = 0; i < count; i++) {
    const at = start + 4 + i * 8, platform = bytes.readUInt16BE(at), encoding = bytes.readUInt16BE(at + 2);
    if (platform !== 0 && !(platform === 3 && [1,10].includes(encoding))) continue;
    const sub = start + bytes.readUInt32BE(at + 4); checkRange(bytes, sub, 2);
    if (sub < start || sub >= start + record.length) invalid();
    const format = bytes.readUInt16BE(sub);
    if (format === 12) {
      checkRange(bytes, sub, 16); const size = bytes.readUInt32BE(sub + 4), groups = bytes.readUInt32BE(sub + 12);
      if (groups > 200000 || size < 16 + groups * 12 || sub + size > start + record.length) invalid();
      for (let j = 0; j < groups; j++) { const g=sub+16+j*12, first=bytes.readUInt32BE(g), last=bytes.readUInt32BE(g+4), glyph=bytes.readUInt32BE(g+8);
        if (first>last || last>0x10ffff) invalid(); if ((examined += last-first+1) > 2200000) invalid(); for (let cp=first+(glyph===0?1:0);cp<=last;cp++) points.add(cp);
      }
    } else if (format === 4) {
      checkRange(bytes, sub, 14); const size=bytes.readUInt16BE(sub+2), segments=bytes.readUInt16BE(sub+6)/2;
      if (!Number.isInteger(segments)||segments>32767||size<16+segments*8||sub+size>start+record.length) invalid();
      const ends=sub+14,starts=ends+segments*2+2,deltas=starts+segments*2,offsets=deltas+segments*2;
      for (let j=0;j<segments;j++) { const first=bytes.readUInt16BE(starts+j*2),last=bytes.readUInt16BE(ends+j*2),delta=bytes.readInt16BE(deltas+j*2),offset=bytes.readUInt16BE(offsets+j*2); if(first>last||(examined+=last-first+1)>2200000)invalid();
        for(let cp=first;cp<=last&&cp!==0xffff;cp++){let glyph;if(offset){const loc=offsets+j*2+offset+2*(cp-first);if(loc+2>sub+size)invalid();glyph=bytes.readUInt16BE(loc);if(glyph)glyph=(glyph+delta)&65535;}else glyph=(cp+delta)&65535;if(glyph)points.add(cp);}
      }
    }
  }
  const sorted=[...points].sort((a,b)=>a-b), ranges=[];
  for(const cp of sorted){const previous=ranges[ranges.length-1];if(previous&&cp===previous[1]+1)previous[1]=cp;else ranges.push([cp,cp]);}
  return ranges;
}

function inspectFont(bytes, tableSet) {
  const os2 = tableSet.records.get('OS/2'), head = tableSet.records.get('head'), fvar = tableSet.records.get('fvar');
  const axes = {};
  if (fvar && fvar.length >= 16) {
    const offset = bytes.readUInt16BE(fvar.offset + 4), count = bytes.readUInt16BE(fvar.offset + 8), size = bytes.readUInt16BE(fvar.offset + 10);
    if (count > 32 || size < 20 || offset + count * size > fvar.length) invalid();
    for (let i = 0; i < count; i++) {
      const at = fvar.offset + offset + i * size, tag = bytes.toString('ascii', at, at + 4);
      axes[tag] = { min: bytes.readInt32BE(at + 4) / 65536, default: bytes.readInt32BE(at + 8) / 65536, max: bytes.readInt32BE(at + 12) / 65536 };
    }
  }
  return {
    coverage: unicodeCoverage(bytes, tableSet.records.get('cmap')),
    axes, style: (os2 && os2.length >= 64 ? bytes.readUInt16BE(os2.offset + 62) & 1 : head && head.length >= 46 ? bytes.readUInt16BE(head.offset + 44) & 2 : false) ? 'italic' : 'normal',
    postscriptNames: nameValues(bytes, tableSet.records.get('name'), 6),
    fsType: os2 && os2.length >= 10 ? bytes.readUInt16BE(os2.offset + 8) : 0,
    defaultWeight: os2 && os2.length >= 6 ? bytes.readUInt16BE(os2.offset + 4) : 400,
    format: tableSet.magic === 0x4f54544f ? 'otf' : 'ttf'
  };
}
function extractFont(input, postscriptName) {
  const bytes = Buffer.from(input);
  if (bytes.length < 12 || bytes.length > MAX_FONT_BYTES) invalid();
  let chosen;
  if (bytes.toString('ascii', 0, 4) === 'ttcf') {
    const count = bytes.readUInt32BE(8);
    if (!count || count > 128) invalid();
    checkRange(bytes, 12, count * 4);
    for (let i = 0; i < count; i++) {
      const candidate = tables(bytes, bytes.readUInt32BE(12 + i * 4));
      if (nameValues(bytes, candidate.records.get('name'), 6).includes(postscriptName)) { chosen = candidate; break; }
      if (count === 1) chosen = candidate;
    }
    if (!chosen) invalid();
  } else { chosen = tables(bytes, 0); return { bytes, ...inspectFont(bytes, chosen), sha256: crypto.createHash('sha256').update(bytes).digest('hex') }; }
  const selected = [...chosen.records.values()].filter(record => record.tag !== 'DSIG').sort((a,b) => a.tag.localeCompare(b.tag));
  const count = selected.length, directoryLength = 12 + count * 16;
  const total = directoryLength + selected.reduce((size, table) => size + ((table.length + 3) & ~3), 0);
  if (total > MAX_FONT_BYTES) invalid();
  const output = Buffer.alloc(total);
  output.writeUInt32BE(chosen.magic, 0); output.writeUInt16BE(count, 4);
  const power = 2 ** Math.floor(Math.log2(count));
  output.writeUInt16BE(power * 16, 6); output.writeUInt16BE(Math.log2(power), 8); output.writeUInt16BE(count * 16 - power * 16, 10);
  let offset = directoryLength, headOffset = null;
  selected.forEach((record, index) => {
    const at = 12 + index * 16;
    const table = Buffer.from(bytes.subarray(record.offset, record.offset + record.length));
    if (record.tag === 'head' && record.length >= 12) { table.writeUInt32BE(0, 8); headOffset = offset; }
    output.write(record.tag, at, 4, 'ascii'); output.writeUInt32BE(checksum(table), at + 4);
    output.writeUInt32BE(offset, at + 8); output.writeUInt32BE(record.length, at + 12);
    table.copy(output, offset); offset += (record.length + 3) & ~3;
  });
  if (headOffset !== null) output.writeUInt32BE((0xb1b0afba - checksum(output)) >>> 0, headOffset + 8);
  const info = inspectFont(output, tables(output, 0));
  return { bytes: output, ...info, sha256: crypto.createHash('sha256').update(output).digest('hex') };
}
module.exports = { MAX_FONT_BYTES, extractFont, checksum, nameValues };

