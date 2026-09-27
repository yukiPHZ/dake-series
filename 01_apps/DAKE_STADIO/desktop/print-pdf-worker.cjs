'use strict';
const zlib = require('node:zlib');
const { isMainThread, parentPort, workerData } = require('node:worker_threads');
function pngRgb(input, width, height) {
  const png = Buffer.from(input);
  if (png.length < 33 || png.subarray(0, 8).toString('hex') !== '89504e470d0a1a0a' || png.readUInt32BE(16) !== width || png.readUInt32BE(20) !== height || png[24] !== 8 || ![2, 6].includes(png[25]) || png[26] || png[27] || png[28]) throw new Error('INVALID_PNG');
  const channels = png[25] === 6 ? 4 : 3, stride = width * channels, chunks = [];
  let ended = false;
  for (let at = 8; at + 12 <= png.length;) {
    const size = png.readUInt32BE(at), type = png.toString('ascii', at + 4, at + 8);
    if (at + size + 12 > png.length) throw new Error('INVALID_PNG');
    if (type === 'IDAT') chunks.push(png.subarray(at + 8, at + 8 + size));
    if (type === 'IEND') { ended = true; break; }
    at += size + 12;
  }
  if (!ended || !chunks.length) throw new Error('INVALID_PNG');
  const expected = (stride + 1) * height;
  const raw = zlib.inflateSync(Buffer.concat(chunks), { maxOutputLength: expected });
  if (raw.length !== expected) throw new Error('INVALID_PNG');
  const rgb = Buffer.allocUnsafe(width * height * 3);
  let previous = Buffer.alloc(stride), current = Buffer.allocUnsafe(stride), target = 0;
  for (let y = 0; y < height; y++) {
    const row = y * (stride + 1), filter = raw[row];
    if (filter > 4) throw new Error('INVALID_PNG');
    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? current[x - channels] : 0, b = previous[x], c = x >= channels ? previous[x - channels] : 0;
      let predict = 0;
      if (filter === 1) predict = a;
      else if (filter === 2) predict = b;
      else if (filter === 3) predict = (a + b) >> 1;
      else if (filter === 4) { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); predict = pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      current[x] = (raw[row + 1 + x] + predict) & 255;
    }
    for (let x = 0; x < stride; x += channels) {
      const alpha = channels === 4 ? current[x + 3] : 255;
      for (let channel = 0; channel < 3; channel++) rgb[target++] = Math.round((current[x + channel] * alpha + 255 * (255 - alpha)) / 255);
    }
    [previous, current] = [current, previous];
  }
  return rgb;
}
function makePdf({bytes,width,height,dpi}) {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || width > 8192 || height > 8192 || width * height > 32000000 || !Number.isFinite(dpi) || dpi < 30 || dpi > 2400) throw new Error('INVALID_SIZE');
  const rgb = pngRgb(bytes,width,height), compressed = zlib.deflateSync(rgb);
  const w = Number((width / dpi * 72).toFixed(8)), h = Number((height / dpi * 72).toFixed(8));
  const stream = Buffer.from('q\n' + w + ' 0 0 ' + h + ' 0 0 cm\n/Im0 Do\nQ\n');
  const objects = [
    Buffer.from('<< /Type /Catalog /Pages 2 0 R >>'),
    Buffer.from('<< /Type /Pages /Kids [3 0 R] /Count 1 >>'),
    Buffer.from('<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ' + w + ' ' + h + '] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>'),
    Buffer.concat([Buffer.from('<< /Type /XObject /Subtype /Image /Width ' + width + ' /Height ' + height + ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length ' + compressed.length + ' >>\nstream\n'),compressed,Buffer.from('\nendstream')]),
    Buffer.concat([Buffer.from('<< /Length ' + stream.length + ' >>\nstream\n'),stream,Buffer.from('endstream')])
  ];
  const parts = [Buffer.from('%PDF-1.7\n%\xE2\xE3\xCF\xD3\n','latin1')], offsets = [0]; let size = parts[0].length;
  objects.forEach((object,index)=>{offsets.push(size);const entry=Buffer.concat([Buffer.from((index+1)+' 0 obj\n'),object,Buffer.from('\nendobj\n')]);parts.push(entry);size+=entry.length});
  const xref = 'xref\n0 '+(objects.length+1)+'\n0000000000 65535 f \n'+offsets.slice(1).map(offset=>String(offset).padStart(10,'0')+' 00000 n \n').join('')+'trailer\n<< /Size '+(objects.length+1)+' /Root 1 0 R >>\nstartxref\n'+size+'\n%%EOF\n';
  parts.push(Buffer.from(xref)); return Buffer.concat(parts);
}
if (!isMainThread) { try { parentPort.postMessage({bytes:makePdf(workerData)}); } catch(error) { parentPort.postMessage({error:error.message}); } }
module.exports = { pngRgb, makePdf };
