/*
 * SnapSize PNG encoder — smallest lossless PNG the browser can make.
 *  ≤256 colours  → palette PNG at 1/2/4/8 bits per pixel (screenshots at 100% are usually here)
 *  opaque        → 24-bit RGB with per-row adaptive filtering
 *  otherwise     → 32-bit RGBA with adaptive filtering
 * Uses the browser's built-in zlib (CompressionStream). Writes 96 DPI so Outlook keeps 1 px = 1 px.
 */
(function (global) {
  "use strict";

  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 255] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function chunk(type, data) {
    const out = new Uint8Array(12 + data.length);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, data.length);
    for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
    out.set(data, 8);
    dv.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
    return out;
  }

  async function zlib(data) {
    const stream = new Blob([data]).stream().pipeThrough(new CompressionStream("deflate"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function paeth(a, b, c) {
    const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
    return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
  }

  /** Adds a filter byte to each row; adaptive = try all five filters and keep the cheapest. */
  function filterRows(raw, h, rowLen, bpp, adaptive) {
    const out = new Uint8Array(h * (rowLen + 1));
    const cand = [0, 1, 2, 3, 4].map(() => new Uint8Array(rowLen));
    let prev = new Uint8Array(rowLen);
    for (let y = 0; y < h; y++) {
      const cur = raw.subarray(y * rowLen, (y + 1) * rowLen);
      let best = 0, bestRow = cur;
      if (adaptive) {
        let bestSum = Infinity;
        for (let f = 0; f < 5; f++) {
          const r = cand[f];
          let sum = 0;
          for (let i = 0; i < rowLen; i++) {
            const a = i >= bpp ? cur[i - bpp] : 0, b = prev[i], c = i >= bpp ? prev[i - bpp] : 0;
            let v;
            switch (f) {
              case 0: v = cur[i]; break;
              case 1: v = cur[i] - a; break;
              case 2: v = cur[i] - b; break;
              case 3: v = cur[i] - ((a + b) >> 1); break;
              default: v = cur[i] - paeth(a, b, c);
            }
            v &= 255;
            r[i] = v;
            sum += v < 128 ? v : 256 - v;
          }
          if (sum < bestSum) { bestSum = sum; best = f; bestRow = r; }
        }
      }
      const o = y * (rowLen + 1);
      out[o] = best;
      out.set(bestRow, o + 1);
      prev = cur;
    }
    return out;
  }

  /** @param {{width:number,height:number,data:Uint8ClampedArray}} img  @returns {Promise<Uint8Array>} */
  async function encode(img) {
    const w = img.width, h = img.height, d = img.data, n = w * h;

    let opaque = true, many = false;
    const map = new Map();
    for (let i = 0; i < n; i++) {
      const o = i * 4;
      if (d[o + 3] !== 255) opaque = false;
      if (!many) {
        const k = ((d[o + 3] << 24) | (d[o] << 16) | (d[o + 1] << 8) | d[o + 2]) >>> 0;
        if (!map.has(k)) {
          if (map.size === 256) many = true;
          else map.set(k, map.size);
        }
      }
    }

    const chunks = [];
    let colorType, depth, idat;

    if (!many) {
      // Palette PNG
      const count = map.size;
      depth = count <= 2 ? 1 : count <= 4 ? 2 : count <= 16 ? 4 : 8;
      colorType = 3;
      const rowLen = Math.ceil((w * depth) / 8);
      const raw = new Uint8Array(rowLen * h);
      const perByte = 8 / depth;
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          const o = (y * w + x) * 4;
          const k = ((d[o + 3] << 24) | (d[o] << 16) | (d[o + 1] << 8) | d[o + 2]) >>> 0;
          const idx = map.get(k);
          const byte = y * rowLen + Math.floor(x / perByte);
          const shift = 8 - depth * ((x % perByte) + 1);
          raw[byte] |= idx << shift;
        }
      }
      const plte = new Uint8Array(count * 3);
      const trns = new Uint8Array(count);
      for (const [k, i] of map) {
        plte[i * 3] = (k >>> 16) & 255;
        plte[i * 3 + 1] = (k >>> 8) & 255;
        plte[i * 3 + 2] = k & 255;
        trns[i] = (k >>> 24) & 255;
      }
      chunks.push(chunk("PLTE", plte));
      if (!opaque) chunks.push(chunk("tRNS", trns));
      idat = await zlib(filterRows(raw, h, rowLen, 1, false));
    } else {
      const bpp = opaque ? 3 : 4;
      colorType = opaque ? 2 : 6;
      depth = 8;
      const rowLen = w * bpp;
      const raw = new Uint8Array(rowLen * h);
      for (let i = 0, j = 0; i < n; i++) {
        raw[j++] = d[i * 4];
        raw[j++] = d[i * 4 + 1];
        raw[j++] = d[i * 4 + 2];
        if (!opaque) raw[j++] = d[i * 4 + 3];
      }
      idat = await zlib(filterRows(raw, h, rowLen, bpp, true));
    }

    const ihdr = new Uint8Array(13);
    const iv = new DataView(ihdr.buffer);
    iv.setUint32(0, w);
    iv.setUint32(4, h);
    ihdr[8] = depth;
    ihdr[9] = colorType;

    const phys = new Uint8Array(9); // 96 DPI = 3780 px/m
    const pv = new DataView(phys.buffer);
    pv.setUint32(0, 3780);
    pv.setUint32(4, 3780);
    phys[8] = 1;

    const parts = [
      new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
      chunk("IHDR", ihdr),
      chunk("pHYs", phys),
      ...chunks,
      chunk("IDAT", idat),
      chunk("IEND", new Uint8Array(0)),
    ];
    const total = parts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(total);
    let off = 0;
    for (const p of parts) { out.set(p, off); off += p.length; }
    return out;
  }

  const api = { encode };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  global.SnapPng = api;
})(typeof window !== "undefined" ? window : globalThis);
