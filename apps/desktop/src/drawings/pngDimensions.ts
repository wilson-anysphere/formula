// Guardrails for image decompression bombs: common image formats can advertise extremely large
// dimensions while the compressed payload remains small. Decoding such images can allocate huge
// bitmaps and hang/crash the renderer.
//
// These limits are shared across:
// - insert/paste flows (so rejected images are never persisted)
// - collaboration hydration (so remote peers can't DoS by publishing huge-dimension PNGs)
// - bitmap decode (defense-in-depth for legacy/persisted images)
export const MAX_PNG_DIMENSION = 10_000;
export const MAX_PNG_PIXELS = 50_000_000;

export type ImageHeaderFormat = "png" | "jpeg" | "gif" | "webp" | "bmp";
export type ImageHeaderDimensions = { format: ImageHeaderFormat; width: number; height: number };

export function readPngDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // PNG signature (8 bytes) + IHDR chunk header (8 bytes) + width/height (8 bytes).
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 24) return null;

  if (
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47 ||
    bytes[4] !== 0x0d ||
    bytes[5] !== 0x0a ||
    bytes[6] !== 0x1a ||
    bytes[7] !== 0x0a
  ) {
    return null;
  }

  // The first chunk should be IHDR: length (4) + type (4) + data...
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) {
    return null;
  }

  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const width = view.getUint32(16, false);
    const height = view.getUint32(20, false);
    if (width === 0 || height === 0) return null;
    return { width, height };
  } catch {
    return null;
  }
}

export function readGifDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // GIF header (6 bytes) + logical screen size (4 bytes).
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 10) return null;
  if (bytes[0] !== 0x47 || bytes[1] !== 0x49 || bytes[2] !== 0x46) return null; // "GIF"
  if (bytes[3] !== 0x38) return null; // "8"
  if (bytes[4] !== 0x37 && bytes[4] !== 0x39) return null; // "7" or "9"
  if (bytes[5] !== 0x61) return null; // "a"
  const width = bytes[6] | (bytes[7] << 8);
  const height = bytes[8] | (bytes[9] << 8);
  if (width === 0 || height === 0) return null;
  return { width, height };
}

export function readJpegDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // JPEG SOI.
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 4) return null;
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;

  // Scan segments until Start of Frame (SOF) or Start of Scan (SOS).
  const len = bytes.byteLength;
  let i = 2;

  const isSofMarker = (marker: number): boolean => {
    // Baseline/progressive DCT + lossless variants.
    // Exclude DHT/DAC/JPG/etc which are not frame headers.
    return (
      marker === 0xc0 ||
      marker === 0xc1 ||
      marker === 0xc2 ||
      marker === 0xc3 ||
      marker === 0xc5 ||
      marker === 0xc6 ||
      marker === 0xc7 ||
      marker === 0xc9 ||
      marker === 0xca ||
      marker === 0xcb ||
      marker === 0xcd ||
      marker === 0xce ||
      marker === 0xcf
    );
  };

  while (i + 1 < len) {
    // Find the next marker prefix (0xFF).
    while (i < len && bytes[i] !== 0xff) i += 1;
    if (i + 1 >= len) break;

    // Skip any padding 0xFF bytes.
    let j = i + 1;
    while (j < len && bytes[j] === 0xff) j += 1;
    if (j >= len) break;

    const marker = bytes[j]!;
    i = j + 1;

    // Standalone markers (no length field).
    if (marker === 0xd8 || marker === 0xd9) {
      // SOI/EOI
      continue;
    }
    if (marker >= 0xd0 && marker <= 0xd7) {
      // RSTn
      continue;
    }
    if (marker === 0x01) {
      // TEM
      continue;
    }
    if (marker === 0xda) {
      // SOS: after this point the stream is entropy-coded; dimensions should have appeared already.
      break;
    }

    // Need a 2-byte segment length.
    if (i + 1 >= len) break;
    const segLen = (bytes[i]! << 8) | bytes[i + 1]!;
    if (segLen < 2) break;
    const segmentEnd = i + segLen;
    if (segmentEnd > len) break;

    if (isSofMarker(marker)) {
      // Segment data starts after the 2 length bytes:
      //  - precision (1)
      //  - height (2)
      //  - width (2)
      if (segLen < 8) return null;
      const data = i + 2;
      const height = (bytes[data + 1]! << 8) | bytes[data + 2]!;
      const width = (bytes[data + 3]! << 8) | bytes[data + 4]!;
      if (width === 0 || height === 0) return null;
      return { width, height };
    }

    // Skip this segment: marker(2 bytes) + length(segLen) where segLen includes the length bytes.
    i = segmentEnd;
  }

  return null;
}

export function readWebpDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // RIFF header (12) + chunk header (8) + minimal VP8/VP8X dimension bytes.
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 30) return null;
  // "RIFF"...."WEBP"
  if (
    bytes[0] !== 0x52 ||
    bytes[1] !== 0x49 ||
    bytes[2] !== 0x46 ||
    bytes[3] !== 0x46 ||
    bytes[8] !== 0x57 ||
    bytes[9] !== 0x45 ||
    bytes[10] !== 0x42 ||
    bytes[11] !== 0x50
  ) {
    return null;
  }

  const chunkType =
    String.fromCharCode(bytes[12]!, bytes[13]!, bytes[14]!, bytes[15]!) as "VP8 " | "VP8L" | "VP8X" | string;

  if (chunkType === "VP8X") {
    // Extended format header:
    // flags (1) + reserved (3) + canvasWidthMinusOne (3 LE) + canvasHeightMinusOne (3 LE)
    const widthMinusOne = bytes[24]! | (bytes[25]! << 8) | (bytes[26]! << 16);
    const heightMinusOne = bytes[27]! | (bytes[28]! << 8) | (bytes[29]! << 16);
    const width = widthMinusOne + 1;
    const height = heightMinusOne + 1;
    if (width <= 0 || height <= 0) return null;
    return { width, height };
  }

  if (chunkType === "VP8 ") {
    // Lossy bitstream: keyframe header stores dimensions at fixed offsets.
    // Frame starts at offset 20 (12 RIFF + 8 chunk header).
    const startCodeOk = bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a;
    if (!startCodeOk) return null;
    const width = (bytes[26]! | (bytes[27]! << 8)) & 0x3fff;
    const height = (bytes[28]! | (bytes[29]! << 8)) & 0x3fff;
    if (width === 0 || height === 0) return null;
    return { width, height };
  }

  if (chunkType === "VP8L") {
    // Lossless: signature 0x2f then 32-bit little-endian bitfield of width/height minus one.
    if (bytes[20] !== 0x2f) return null;
    const bits =
      (bytes[21]! | (bytes[22]! << 8) | (bytes[23]! << 16) | (bytes[24]! << 24)) >>> 0;
    const width = 1 + (bits & 0x3fff);
    const height = 1 + ((bits >>> 14) & 0x3fff);
    if (width === 0 || height === 0) return null;
    return { width, height };
  }

  return null;
}

export function readBmpDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // BMP file header (14 bytes) + DIB header size (4 bytes) + width/height (>=4 bytes each)
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 26) return null;
  // "BM"
  if (bytes[0] !== 0x42 || bytes[1] !== 0x4d) return null;

  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const dibSize = view.getUint32(14, true);

    if (dibSize === 12) {
      // BITMAPCOREHEADER: 16-bit width/height (unsigned).
      const width = view.getUint16(18, true);
      const height = view.getUint16(20, true);
      if (width === 0 || height === 0) return null;
      return { width, height };
    }

    if (dibSize >= 40) {
      // BITMAPINFOHEADER and later: 32-bit signed width/height (height may be negative for top-down).
      const widthRaw = view.getInt32(18, true);
      const heightRaw = view.getInt32(22, true);
      const width = Math.abs(widthRaw);
      const height = Math.abs(heightRaw);
      if (width === 0 || height === 0) return null;
      return { width, height };
    }
  } catch {
    return null;
  }

  return null;
}

export function readImageDimensions(bytes: Uint8Array): ImageHeaderDimensions | null {
  const png = readPngDimensions(bytes);
  if (png) return { format: "png", ...png };
  const gif = readGifDimensions(bytes);
  if (gif) return { format: "gif", ...gif };
  const jpeg = readJpegDimensions(bytes);
  if (jpeg) return { format: "jpeg", ...jpeg };
  const webp = readWebpDimensions(bytes);
  if (webp) return { format: "webp", ...webp };
  const bmp = readBmpDimensions(bytes);
  if (bmp) return { format: "bmp", ...bmp };
  return null;
}
