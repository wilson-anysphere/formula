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

export type ImageHeaderFormat = "png" | "jpeg" | "gif" | "webp" | "bmp" | "svg";
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

export function readSvgDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 4) return null;

  const len = bytes.byteLength;

  const decodeUtf16 = (view: Uint8Array, littleEndian: boolean): string | null => {
    const len = view.byteLength - (view.byteLength % 2);
    if (len <= 0) return null;
    let bytesToDecode = view.subarray(0, len);

    if (!littleEndian) {
      // Convert BE -> LE so we can decode via TextDecoder("utf-16le") in environments that
      // do not support "utf-16be".
      const swapped = new Uint8Array(len);
      for (let i = 0; i < len; i += 2) {
        swapped[i] = bytesToDecode[i + 1]!;
        swapped[i + 1] = bytesToDecode[i]!;
      }
      bytesToDecode = swapped;
    }

    try {
      if (typeof TextDecoder !== "undefined") {
        return new TextDecoder("utf-16le", { fatal: false }).decode(bytesToDecode);
      }
    } catch {
      // Fall through to manual decoding.
    }

    try {
      let out = "";
      for (let i = 0; i < bytesToDecode.length; i += 2) {
        out += String.fromCharCode(bytesToDecode[i]! | (bytesToDecode[i + 1]! << 8));
      }
      return out;
    } catch {
      return null;
    }
  };

  const decodeUtf8 = (view: Uint8Array): string | null => {
    try {
      if (typeof TextDecoder !== "undefined") {
        return new TextDecoder("utf-8", { fatal: false }).decode(view);
      }
    } catch {
      // Fall through to manual decoding.
    }
    try {
      let out = "";
      for (let i = 0; i < view.length; i += 1) out += String.fromCharCode(view[i]!);
      return out;
    } catch {
      return null;
    }
  };

  const isWhitespaceByte = (b: number): boolean => b === 0x20 || b === 0x09 || b === 0x0a || b === 0x0d;

  // Fast-path: SVGs are text-based. If we don't see `<` (optionally after BOM/whitespace) or a UTF-16 BOM,
  // don't spend time scanning/decoding.
  const utf16Bom = len >= 2 && ((bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0xff && bytes[1] === 0xfe));
  const utf16LittleEndian = utf16Bom ? bytes[0] === 0xff && bytes[1] === 0xfe : false;

  const readUtf16CodeUnit = (idx: number): number =>
    utf16LittleEndian ? bytes[idx]! | (bytes[idx + 1]! << 8) : (bytes[idx]! << 8) | bytes[idx + 1]!;

  if (utf16Bom) {
    let idx = 2;
    while (idx + 1 < len) {
      const cu = readUtf16CodeUnit(idx);
      if (cu === 0x20 || cu === 0x09 || cu === 0x0a || cu === 0x0d) {
        idx += 2;
        continue;
      }
      break;
    }
    if (idx + 1 >= len) return null;
    if (readUtf16CodeUnit(idx) !== 0x3c) return null; // '<'
  } else {
    let idx = 0;
    if (len >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) idx = 3; // UTF-8 BOM
    while (idx < len && isWhitespaceByte(bytes[idx]!)) idx += 1;
    if (idx >= len) return null;
    if (bytes[idx] !== 0x3c) return null; // '<'
  }

  const MAX_TAG_BYTES = 32 * 1024;

  const findSvgTagBytesUtf8 = (): Uint8Array | null => {
    const asciiLower = (b: number): number => (b >= 0x41 && b <= 0x5a ? b | 0x20 : b);

    const isSvgName = (start: number, end: number): boolean => {
      const nameLen = end - start;
      if (nameLen === 3) {
        return (
          asciiLower(bytes[start]!) === 0x73 &&
          asciiLower(bytes[start + 1]!) === 0x76 &&
          asciiLower(bytes[start + 2]!) === 0x67
        );
      }
      if (nameLen > 4 && bytes[end - 4] === 0x3a) {
        return (
          asciiLower(bytes[end - 3]!) === 0x73 &&
          asciiLower(bytes[end - 2]!) === 0x76 &&
          asciiLower(bytes[end - 1]!) === 0x67
        );
      }
      return false;
    };

    const findTagEnd = (start: number): number => {
      const limit = Math.min(len, start + MAX_TAG_BYTES);
      let quote = 0;
      for (let i = start + 1; i < limit; i += 1) {
        const b = bytes[i]!;
        if (quote) {
          if (b === quote) quote = 0;
          continue;
        }
        if (b === 0x22 || b === 0x27) {
          quote = b;
          continue;
        }
        if (b === 0x3e) return i + 1; // '>'
      }
      return -1;
    };

    let i = 0;
    if (len >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) i = 3;

    while (i < len) {
      while (i < len && bytes[i] !== 0x3c) i += 1; // '<'
      if (i >= len || i + 1 >= len) break;

      const next = bytes[i + 1]!;

      if (next === 0x21) {
        // Declarations: comments, CDATA, doctype.
        if (i + 3 < len && bytes[i + 2] === 0x2d && bytes[i + 3] === 0x2d) {
          // <!-- ... -->
          let j = i + 4;
          while (j + 2 < len) {
            if (bytes[j] === 0x2d && bytes[j + 1] === 0x2d && bytes[j + 2] === 0x3e) {
              i = j + 3;
              break;
            }
            j += 1;
          }
          if (j + 2 >= len) return null;
          continue;
        }

        // <![CDATA[ ... ]]>
        const cdataStart =
          i + 8 < len &&
          bytes[i + 2] === 0x5b &&
          bytes[i + 3] === 0x43 &&
          bytes[i + 4] === 0x44 &&
          bytes[i + 5] === 0x41 &&
          bytes[i + 6] === 0x54 &&
          bytes[i + 7] === 0x41 &&
          bytes[i + 8] === 0x5b;
        if (cdataStart) {
          let j = i + 9;
          while (j + 2 < len) {
            if (bytes[j] === 0x5d && bytes[j + 1] === 0x5d && bytes[j + 2] === 0x3e) {
              i = j + 3;
              break;
            }
            j += 1;
          }
          if (j + 2 >= len) return null;
          continue;
        }

        // Other declaration (e.g. <!DOCTYPE ...>): skip until `>` not in quotes and not inside `[ ... ]`.
        let j = i + 2;
        let quote = 0;
        let bracketDepth = 0;
        while (j < len) {
          const b = bytes[j]!;
          if (quote) {
            if (b === quote) quote = 0;
          } else if (b === 0x22 || b === 0x27) {
            quote = b;
          } else if (b === 0x5b) {
            bracketDepth += 1;
          } else if (b === 0x5d && bracketDepth > 0) {
            bracketDepth -= 1;
          } else if (b === 0x3e && bracketDepth === 0) {
            i = j + 1;
            break;
          }
          j += 1;
        }
        if (j >= len) return null;
        continue;
      }

      if (next === 0x3f) {
        // Processing instruction: <? ... ?>
        let j = i + 2;
        while (j + 1 < len) {
          if (bytes[j] === 0x3f && bytes[j + 1] === 0x3e) {
            i = j + 2;
            break;
          }
          j += 1;
        }
        if (j + 1 >= len) return null;
        continue;
      }

      if (next === 0x2f) {
        // End tag: </...>
        const end = findTagEnd(i);
        if (end === -1) return null;
        i = end;
        continue;
      }

      // Regular start tag.
      const nameStart = i + 1;
      let nameEnd = nameStart;
      const maxNameEnd = Math.min(len, nameStart + 64);
      while (nameEnd < maxNameEnd) {
        const b = bytes[nameEnd]!;
        if (isWhitespaceByte(b) || b === 0x2f || b === 0x3e || b === 0x3f) break;
        nameEnd += 1;
      }

      const end = findTagEnd(i);
      if (end === -1) return null;

      if (nameEnd > nameStart && isSvgName(nameStart, nameEnd)) {
        return bytes.subarray(i, end);
      }

      i = end;
    }

    return null;
  };

  const findSvgTagBytesUtf16 = (): Uint8Array | null => {
    const asciiLower = (cu: number): number => (cu >= 0x41 && cu <= 0x5a ? cu | 0x20 : cu);
    const isWhitespaceCu = (cu: number): boolean => cu === 0x20 || cu === 0x09 || cu === 0x0a || cu === 0x0d;

    const isSvgName = (start: number, end: number): boolean => {
      const nameLen = (end - start) / 2;
      if (nameLen === 3) {
        return (
          asciiLower(readUtf16CodeUnit(start)) === 0x73 &&
          asciiLower(readUtf16CodeUnit(start + 2)) === 0x76 &&
          asciiLower(readUtf16CodeUnit(start + 4)) === 0x67
        );
      }
      if (nameLen > 4 && readUtf16CodeUnit(end - 8) === 0x3a) {
        return (
          asciiLower(readUtf16CodeUnit(end - 6)) === 0x73 &&
          asciiLower(readUtf16CodeUnit(end - 4)) === 0x76 &&
          asciiLower(readUtf16CodeUnit(end - 2)) === 0x67
        );
      }
      return false;
    };

    const findTagEnd = (start: number): number => {
      const limit = Math.min(len - (len % 2), start + MAX_TAG_BYTES);
      let quote = 0;
      for (let i = start + 2; i + 1 < limit; i += 2) {
        const cu = readUtf16CodeUnit(i);
        if (quote) {
          if (cu === quote) quote = 0;
          continue;
        }
        if (cu === 0x22 || cu === 0x27) {
          quote = cu;
          continue;
        }
        if (cu === 0x3e) return i + 2; // '>'
      }
      return -1;
    };

    let i = 2; // skip BOM
    while (i + 1 < len) {
      while (i + 1 < len && readUtf16CodeUnit(i) !== 0x3c) i += 2; // '<'
      if (i + 1 >= len || i + 3 >= len) break;

      const next = readUtf16CodeUnit(i + 2);

      if (next === 0x21) {
        // Declarations: comments, CDATA, doctype.
        if (i + 7 < len && readUtf16CodeUnit(i + 4) === 0x2d && readUtf16CodeUnit(i + 6) === 0x2d) {
          // <!-- ... -->
          let j = i + 8;
          while (j + 5 < len) {
            if (
              readUtf16CodeUnit(j) === 0x2d &&
              readUtf16CodeUnit(j + 2) === 0x2d &&
              readUtf16CodeUnit(j + 4) === 0x3e
            ) {
              i = j + 6;
              break;
            }
            j += 2;
          }
          if (j + 5 >= len) return null;
          continue;
        }

        // <![CDATA[ ... ]]>
        const cdataStart =
          i + 17 < len &&
          readUtf16CodeUnit(i + 4) === 0x5b &&
          readUtf16CodeUnit(i + 6) === 0x43 &&
          readUtf16CodeUnit(i + 8) === 0x44 &&
          readUtf16CodeUnit(i + 10) === 0x41 &&
          readUtf16CodeUnit(i + 12) === 0x54 &&
          readUtf16CodeUnit(i + 14) === 0x41 &&
          readUtf16CodeUnit(i + 16) === 0x5b;
        if (cdataStart) {
          let j = i + 18;
          while (j + 5 < len) {
            if (
              readUtf16CodeUnit(j) === 0x5d &&
              readUtf16CodeUnit(j + 2) === 0x5d &&
              readUtf16CodeUnit(j + 4) === 0x3e
            ) {
              i = j + 6;
              break;
            }
            j += 2;
          }
          if (j + 5 >= len) return null;
          continue;
        }

        // Other declaration: skip until `>` not in quotes and not inside `[ ... ]`.
        let j = i + 4;
        let quote = 0;
        let bracketDepth = 0;
        while (j + 1 < len) {
          const cu = readUtf16CodeUnit(j);
          if (quote) {
            if (cu === quote) quote = 0;
          } else if (cu === 0x22 || cu === 0x27) {
            quote = cu;
          } else if (cu === 0x5b) {
            bracketDepth += 1;
          } else if (cu === 0x5d && bracketDepth > 0) {
            bracketDepth -= 1;
          } else if (cu === 0x3e && bracketDepth === 0) {
            i = j + 2;
            break;
          }
          j += 2;
        }
        if (j + 1 >= len) return null;
        continue;
      }

      if (next === 0x3f) {
        // Processing instruction: <? ... ?>
        let j = i + 4;
        while (j + 3 < len) {
          if (readUtf16CodeUnit(j) === 0x3f && readUtf16CodeUnit(j + 2) === 0x3e) {
            i = j + 4;
            break;
          }
          j += 2;
        }
        if (j + 3 >= len) return null;
        continue;
      }

      if (next === 0x2f) {
        // End tag.
        const end = findTagEnd(i);
        if (end === -1) return null;
        i = end;
        continue;
      }

      // Regular start tag.
      const nameStart = i + 2;
      let nameEnd = nameStart;
      const maxNameEnd = Math.min(len - (len % 2), nameStart + 128);
      while (nameEnd + 1 < maxNameEnd) {
        const cu = readUtf16CodeUnit(nameEnd);
        if (isWhitespaceCu(cu) || cu === 0x2f || cu === 0x3e || cu === 0x3f) break;
        nameEnd += 2;
      }

      const end = findTagEnd(i);
      if (end === -1) return null;

      if (nameEnd > nameStart && isSvgName(nameStart, nameEnd)) {
        return bytes.subarray(i, end);
      }

      i = end;
    }

    return null;
  };

  const tagBytes = utf16Bom ? findSvgTagBytesUtf16() : findSvgTagBytesUtf8();
  if (!tagBytes) return null;

  const tag = utf16Bom ? decodeUtf16(tagBytes, utf16LittleEndian) : decodeUtf8(tagBytes);
  if (!tag) return null;

  const getAttr = (name: string): string | null => {
    // Match full attribute names only (avoid matching `data-width` / `stroke-width` / etc).
    // Attributes in XML must be separated by whitespace, so we can safely require it here.
    const re = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])(.*?)\\1`, "i");
    const match = re.exec(tag);
    return match ? String(match[2] ?? "").trim() : null;
  };

  const parseLength = (raw: string | null): number | null => {
    if (!raw) return null;
    let normalized = raw.trim();
    // Support simple `calc(<length>)` wrappers used in some SVGs.
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const calc = /^\s*calc\(\s*(.+)\s*\)\s*$/i.exec(normalized);
      if (!calc) break;
      normalized = String(calc[1] ?? "").trim();
    }
    const m =
      /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\s*([a-z%]*)\s*$/i.exec(normalized);
    if (!m) return null;
    const value = Number(m[1]);
    if (!Number.isFinite(value) || value <= 0) return null;
    const unit = String(m[2] ?? "").toLowerCase();
    switch (unit) {
      case "":
      case "px":
        return value;
      case "%":
        return null;
      case "pt":
        return (value * 96) / 72;
      case "pc":
        return value * 16;
      case "in":
        return value * 96;
      case "cm":
        return (value * 96) / 2.54;
      case "mm":
        return (value * 96) / 25.4;
      case "q":
        return (value * 96) / 101.6;
      case "em":
        return value * 16;
      case "ex":
        return value * 8;
      default:
        // Unknown units: treat as user units (best-effort). This is more conservative than returning
        // null because it prevents bypassing dimension guards via unusual unit strings.
        return value;
    }
  };

  const widthAttr = parseLength(getAttr("width"));
  const heightAttr = parseLength(getAttr("height"));

  let width = widthAttr;
  let height = heightAttr;

  if (width == null || height == null) {
    const style = getAttr("style");
    if (style) {
      const getStyleProp = (prop: "width" | "height"): number | null => {
        const re = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, "i");
        const match = re.exec(style);
        if (!match) return null;
        const raw = String(match[1] ?? "").replace(/!important\\b/i, "").trim();
        return parseLength(raw);
      };
      if (width == null) width = getStyleProp("width");
      if (height == null) height = getStyleProp("height");
    }
  }

  if (width == null || height == null) {
    const viewBox = getAttr("viewBox");
    if (viewBox) {
      const parts = viewBox
        .trim()
        .split(/[\s,]+/)
        .map((part) => part.trim())
        .filter(Boolean);
      if (parts.length >= 4) {
        const vbWidth = Number(parts[2]);
        const vbHeight = Number(parts[3]);
        if (width == null && Number.isFinite(vbWidth) && vbWidth > 0) width = vbWidth;
        if (height == null && Number.isFinite(vbHeight) && vbHeight > 0) height = vbHeight;
      }
    }
  }

  if (width == null || height == null) return null;
  if (!Number.isFinite(width) || !Number.isFinite(height)) return null;
  if (width <= 0 || height <= 0) return null;
  return { width, height };
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
  const svg = readSvgDimensions(bytes);
  if (svg) return { format: "svg", ...svg };
  return null;
}
