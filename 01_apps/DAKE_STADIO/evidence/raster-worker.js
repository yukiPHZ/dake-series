/* Pixel formulas match Fabric.js 7.4 CPU filters (MIT, see THIRD_PARTY_NOTICES).
 * Each pass writes into Uint8ClampedArray, preserving Fabric's rounding order.
 * Only one transferred bitmap is processed at a time by StudioEngine. */
self.onmessage = ({ data: { id, bitmap, filters } }) => {
  try {
    const width = bitmap.width; const height = bitmap.height;
    if (width < 1 || height < 1 || width > 8192 || height > 8192 || width * height > 32_000_000) throw new Error('DOCUMENT_TOO_LARGE');
    const canvas = new OffscreenCanvas(width, height);
    const context = canvas.getContext('2d', { willReadFrequently: true });
    context.drawImage(bitmap, 0, 0); bitmap.close();
    const imageData = context.getImageData(0, 0, width, height);
    const pixels = imageData.data;
    for (const filter of filters) {
      if (filter.type === 'Brightness') {
        const amount = Math.round(filter.brightness * 255);
        for (let i = 0; i < pixels.length; i += 4) { pixels[i] += amount; pixels[i + 1] += amount; pixels[i + 2] += amount; }
      } else if (filter.type === 'Contrast') {
        const amount = Math.floor(filter.contrast * 255);
        const factor = (259 * (amount + 255)) / (255 * (259 - amount));
        for (let i = 0; i < pixels.length; i += 4) {
          pixels[i] = factor * (pixels[i] - 128) + 128;
          pixels[i + 1] = factor * (pixels[i + 1] - 128) + 128;
          pixels[i + 2] = factor * (pixels[i + 2] - 128) + 128;
        }
      } else if (filter.type === 'Saturation') {
        const adjust = -filter.saturation;
        for (let i = 0; i < pixels.length; i += 4) {
          const r = pixels[i]; const g = pixels[i + 1]; const b = pixels[i + 2]; const max = Math.max(r, g, b);
          pixels[i] += max !== r ? (max - r) * adjust : 0;
          pixels[i + 1] += max !== g ? (max - g) * adjust : 0;
          pixels[i + 2] += max !== b ? (max - b) * adjust : 0;
        }
      } else if (filter.type === 'Grayscale') {
        for (let i = 0; i < pixels.length; i += 4) {
          const r = pixels[i]; const g = pixels[i + 1]; const b = pixels[i + 2];
          const value = filter.mode === 'lightness' ? (Math.min(r, g, b) + Math.max(r, g, b)) / 2 :
            filter.mode === 'luminosity' ? r * 0.21 + g * 0.72 + b * 0.07 : (r + g + b) / 3;
          pixels[i] = pixels[i + 1] = pixels[i + 2] = value;
        }
      } else throw new Error('INVALID_DOCUMENT');
    }
    context.putImageData(imageData, 0, 0);
    const result = canvas.transferToImageBitmap();
    self.postMessage({ id, bitmap: result }, [result]);
  } catch (error) {
    bitmap?.close();
    self.postMessage({ id, error: error.message || 'INVALID_IMAGE' });
  }
};
