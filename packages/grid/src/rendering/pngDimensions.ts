// Guardrails for image decompression bombs: common image formats can advertise extremely large
// dimensions while the compressed payload remains small. Decoding such images can allocate huge
// bitmaps and hang/crash the renderer.
//
// These limits are shared with the desktop renderer (drawings) for consistency.
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

  type TagBounds = { start: number; end: number };

  const findSvgTagBytesUtf8 = (): TagBounds | null => {
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
        return { start: i, end };
      }

      i = end;
    }

    return null;
  };

  const findSvgTagBytesUtf16 = (): TagBounds | null => {
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
        return { start: i, end };
      }

      i = end;
    }

    return null;
  };

  const tagBounds = utf16Bom ? findSvgTagBytesUtf16() : findSvgTagBytesUtf8();
  if (!tagBounds) return null;
  const tagBytes = bytes.subarray(tagBounds.start, tagBounds.end);
  const tagEnd = tagBounds.end;

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
    let hadCalcWrapper = false;
    // Support simple `calc(...)` wrappers used in some SVGs.
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const calc = /^\s*calc\(\s*(.+)\s*\)\s*$/i.exec(normalized);
      if (!calc) break;
      hadCalcWrapper = true;
      normalized = String(calc[1] ?? "").trim();
    }

    const convertToPx = (value: number, unitRaw: string): number | null => {
      if (!Number.isFinite(value)) return null;
      const unit = String(unitRaw ?? "").toLowerCase();
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

    const direct =
      /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\s*([a-z%]*)\s*$/i.exec(normalized);
    if (direct) {
      const value = Number(direct[1]);
      if (!Number.isFinite(value) || value <= 0) return null;
      const px = convertToPx(value, String(direct[2] ?? ""));
      return px && px > 0 ? px : null;
    }

    if (!hadCalcWrapper) return null;

    // Support a limited subset of calc expressions so `calc(1px - -10000px)` or `calc(10001px * 1)`
    // cannot bypass guards. We intentionally support only simple arithmetic and disallow nested
    // parentheses.
    const parseCalcExpression = (expr: string): number | null => {
      const source = expr.trim();
      if (!source) return null;
      if (/[()]/.test(source)) return null;

      type Token =
        | { t: "op"; v: "+" | "-" | "*" | "/" }
        | { t: "number"; n: number }
        | { t: "length"; px: number };

      const tokens: Token[] = [];

      const numberRe = /^(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?/i;
      let idx = 0;
      while (idx < source.length) {
        while (idx < source.length && /\s/.test(source[idx]!)) idx += 1;
        if (idx >= source.length) break;
        const ch = source[idx]!;
        if (ch === "+" || ch === "-" || ch === "*" || ch === "/") {
          tokens.push({ t: "op", v: ch });
          idx += 1;
          continue;
        }

        const match = numberRe.exec(source.slice(idx));
        if (!match) return null;
        const literal = match[0];
        const n = Number(literal);
        if (!Number.isFinite(n)) return null;
        idx += literal.length;

        const unitStart = idx;
        while (idx < source.length && /[a-z%]/i.test(source[idx]!)) idx += 1;
        const unit = source.slice(unitStart, idx);

        if (!unit) {
          tokens.push({ t: "number", n });
          continue;
        }

        const px = convertToPx(n, unit);
        if (px == null) return null;
        tokens.push({ t: "length", px });
      }

      type Value = { k: "number"; n: number } | { k: "length"; px: number };
      let pos = 0;

      const peekOp = (): ("+" | "-" | "*" | "/") | null => {
        const tok = tokens[pos];
        return tok && tok.t === "op" ? tok.v : null;
      };

      const parsePrimary = (): Value | null => {
        const tok = tokens[pos];
        if (!tok) return null;
        if (tok.t === "number") {
          pos += 1;
          return { k: "number", n: tok.n };
        }
        if (tok.t === "length") {
          pos += 1;
          return { k: "length", px: tok.px };
        }
        return null;
      };

      const parseUnary = (): Value | null => {
        let sign = 1;
        while (true) {
          const op = peekOp();
          if (op === "+") {
            pos += 1;
            continue;
          }
          if (op === "-") {
            pos += 1;
            sign = -sign;
            continue;
          }
          break;
        }

        const value = parsePrimary();
        if (!value) return null;
        if (sign === 1) return value;
        return value.k === "length" ? { k: "length", px: value.px * sign } : { k: "number", n: value.n * sign };
      };

      const parseProduct = (): Value | null => {
        let left = parseUnary();
        if (!left) return null;
        while (true) {
          const op = peekOp();
          if (op !== "*" && op !== "/") break;
          pos += 1;
          const right = parseUnary();
          if (!right) return null;

          if (op === "*") {
            if (left.k === "length" && right.k === "number") {
              left = { k: "length", px: left.px * right.n };
            } else if (left.k === "number" && right.k === "length") {
              left = { k: "length", px: left.n * right.px };
            } else if (left.k === "number" && right.k === "number") {
              left = { k: "number", n: left.n * right.n };
            } else {
              return null;
            }
          } else {
            // Division
            if (right.k !== "number") return null;
            if (right.n === 0) return null;
            if (left.k === "length") {
              left = { k: "length", px: left.px / right.n };
            } else {
              left = { k: "number", n: left.n / right.n };
            }
          }

          if (!Number.isFinite(left.k === "length" ? left.px : left.n)) return null;
        }
        return left;
      };

      const parseSum = (): Value | null => {
        let left = parseProduct();
        if (!left) return null;
        while (true) {
          const op = peekOp();
          if (op !== "+" && op !== "-") break;
          pos += 1;
          const right = parseProduct();
          if (!right) return null;
          if (left.k !== right.k) return null;
          if (left.k === "length") {
            left = { k: "length", px: op === "+" ? left.px + right.px : left.px - right.px };
          } else {
            left = { k: "number", n: op === "+" ? left.n + right.n : left.n - right.n };
          }
          if (!Number.isFinite(left.k === "length" ? left.px : left.n)) return null;
        }
        return left;
      };

      const result = parseSum();
      if (!result) return null;
      if (pos !== tokens.length) return null;
      if (result.k !== "length") return null;
      return result.px > 0 ? result.px : null;
    };

    return parseCalcExpression(normalized);
  };

  const widthAttr = parseLength(getAttr("width"));
  const heightAttr = parseLength(getAttr("height"));

  let width = widthAttr;
  let height = heightAttr;

  if (width == null || height == null) {
    const style = getAttr("style");
    if (style) {
      const vars = (() => {
        const map = new Map<string, string>();
        for (const rawDecl of style.split(";")) {
          const decl = rawDecl.trim();
          if (!decl) continue;
          const colon = decl.indexOf(":");
          if (colon === -1) continue;
          const name = decl.slice(0, colon).trim();
          if (!name.startsWith("--")) continue;
          const value = decl.slice(colon + 1).replace(/!important\\b/i, "").trim();
          if (!value) continue;
          map.set(name, value);
        }
        return map;
      })();

      const resolveCssVars = (value: string): string => {
        let current = value;
        for (let iter = 0; iter < 10; iter += 1) {
          const lower = current.toLowerCase();
          const first = lower.indexOf("var(");
          if (first === -1) break;
          let out = "";
          let idx = 0;
          while (idx < current.length) {
            const start = lower.indexOf("var(", idx);
            if (start === -1) {
              out += current.slice(idx);
              break;
            }
            out += current.slice(idx, start);
            let i = start + 4;
            let depth = 1;
            while (i < current.length && depth > 0) {
              const ch = current[i]!;
              if (ch === "(") depth += 1;
              else if (ch === ")") depth -= 1;
              i += 1;
            }
            if (depth !== 0) {
              // Unbalanced; keep the rest as-is.
              out += current.slice(start);
              idx = current.length;
              break;
            }
            const inner = current.slice(start + 4, i - 1).trim();
            // Split "name, fallback" by the first comma at depth 0.
            let namePart = inner;
            let fallback: string | null = null;
            {
              let j = 0;
              let paren = 0;
              for (; j < inner.length; j += 1) {
                const ch = inner[j]!;
                if (ch === "(") paren += 1;
                else if (ch === ")") paren = Math.max(0, paren - 1);
                else if (ch === "," && paren === 0) break;
              }
              if (j < inner.length) {
                namePart = inner.slice(0, j);
                fallback = inner.slice(j + 1);
              }
            }

            const name = namePart.trim();
            const replacement = vars.get(name) ?? (fallback != null ? fallback.trim() : "");
            out += replacement;
            idx = i;
          }
          if (out === current) break;
          current = out;
        }
        return current;
      };

      const getStyleProp = (prop: "width" | "height"): number | null => {
        const re = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, "i");
        const match = re.exec(style);
        if (!match) return null;
        const raw = String(match[1] ?? "").replace(/!important\\b/i, "").trim();
        return parseLength(resolveCssVars(raw));
      };
      if (width == null) width = getStyleProp("width");
      if (height == null) height = getStyleProp("height");
    }
  }

  if (width == null || height == null) {
    const svgId = getAttr("id");
    const svgClasses = (() => {
      const raw = getAttr("class");
      if (!raw) return [] as string[];
      return raw
        .split(/\s+/)
        .map((part) => part.trim())
        .filter(Boolean);
    })();

    type Candidate = { spec: number; order: number; value: string };

    const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    const selectorMatchesRoot = (selector: string): boolean => {
      const sel = selector.trim();
      if (!sel) return false;
      // Ignore combinators/descendant selectors; we only handle simple selectors that target the root element directly.
      if (/[\s>+~]/.test(sel)) return false;

      const typeMatch = /^[a-z][a-z0-9_-]*/i.exec(sel);
      if (typeMatch && typeMatch[0].toLowerCase() !== "svg") {
        // `rect:root { ... }` should not match.
        return false;
      }

      const lower = sel.toLowerCase();
      const hasSvgType = /^svg(?:$|[.#:\[])/i.test(sel);
      const hasRootPseudo = lower.includes(":root");
      const hasId =
        svgId != null && svgId.length > 0
          ? new RegExp(`#${escapeRegExp(svgId.toLowerCase())}(?:$|[.#:\\[])`).test(lower)
          : false;
      const hasClass =
        svgClasses.length > 0
          ? svgClasses.some((cls) => new RegExp(`\\.${escapeRegExp(cls.toLowerCase())}(?:$|[.#:\\[])`).test(lower))
          : false;

      return hasSvgType || hasRootPseudo || hasId || hasClass;
    };

    const computeSpecificity = (selector: string): number => {
      // We only operate on simple selectors (no combinators), so a lightweight specificity estimate is sufficient:
      // id selectors (#) > class/pseudo/attribute selectors (./:/[) > type selectors (svg).
      let ids = 0;
      let classes = 0;
      let tags = 0;

      for (let i = 0; i < selector.length; i += 1) {
        const ch = selector[i]!;
        if (ch === "#") ids += 1;
        else if (ch === "." || ch === "[") classes += 1;
        else if (ch === ":") {
          // Treat pseudo-classes (but not pseudo-elements) as class specificity.
          if (selector[i + 1] !== ":") classes += 1;
        }
      }

      if (/^svg(?:$|[.#:\[])/i.test(selector)) tags = 1;
      return ids * 100 + classes * 10 + tags;
    };

    const updateCandidate = (current: Candidate | null, next: Candidate): Candidate => {
      if (!current) return next;
      if (next.spec > current.spec) return next;
      if (next.spec === current.spec && next.order > current.order) return next;
      return current;
    };

    const stripXmlComments = (input: string): string => {
      let out = "";
      let idx = 0;
      while (idx < input.length) {
        const start = input.indexOf("<!--", idx);
        if (start === -1) {
          out += input.slice(idx);
          break;
        }
        out += input.slice(idx, start);
        const end = input.indexOf("-->", start + 4);
        if (end === -1) break;
        idx = end + 3;
      }
      return out;
    };

    const stripCssComments = (input: string): string => {
      let out = "";
      let idx = 0;
      while (idx < input.length) {
        const start = input.indexOf("/*", idx);
        if (start === -1) {
          out += input.slice(idx);
          break;
        }
        out += input.slice(idx, start);
        const end = input.indexOf("*/", start + 2);
        if (end === -1) break;
        idx = end + 2;
      }
      return out;
    };

    const resolveCssVars = (value: string, vars: Map<string, string>): string => {
      let current = value;
      for (let iter = 0; iter < 10; iter += 1) {
        const lower = current.toLowerCase();
        const first = lower.indexOf("var(");
        if (first === -1) break;
        let out = "";
        let idx = 0;
        while (idx < current.length) {
          const start = lower.indexOf("var(", idx);
          if (start === -1) {
            out += current.slice(idx);
            break;
          }
          out += current.slice(idx, start);
          let i = start + 4;
          let depth = 1;
          while (i < current.length && depth > 0) {
            const ch = current[i]!;
            if (ch === "(") depth += 1;
            else if (ch === ")") depth -= 1;
            i += 1;
          }
          if (depth !== 0) {
            // Unbalanced; keep the rest as-is.
            out += current.slice(start);
            idx = current.length;
            break;
          }
          const inner = current.slice(start + 4, i - 1).trim();
          // Split "name, fallback" by the first comma at depth 0.
          let namePart = inner;
          let fallback: string | null = null;
          {
            let j = 0;
            let paren = 0;
            for (; j < inner.length; j += 1) {
              const ch = inner[j]!;
              if (ch === "(") paren += 1;
              else if (ch === ")") paren = Math.max(0, paren - 1);
              else if (ch === "," && paren === 0) break;
            }
            if (j < inner.length) {
              namePart = inner.slice(0, j);
              fallback = inner.slice(j + 1);
            }
          }

          const name = namePart.trim();
          const replacement = vars.get(name) ?? (fallback != null ? fallback.trim() : "");
          out += replacement;
          idx = i;
        }
        if (out === current) break;
        current = out;
      }
      return current;
    };

    const varCandidates = new Map<string, Candidate>();
    let bestWidth: Candidate | null = null;
    let bestHeight: Candidate | null = null;
    let order = 0;

    const ingestCss = (cssRaw: string): void => {
      let css = cssRaw;
      // Common SVG pattern: wrap CSS in CDATA so `<` characters remain valid XML.
      css = css.split("<![CDATA[").join("").split("]]>").join("");
      // Some SVGs still use XML comments inside `<style>` tags.
      css = stripXmlComments(css);
      css = stripCssComments(css);

      const MAX_CSS_DEPTH = 5;
      const MAX_CSS_RULES = 500;
      let rulesSeen = 0;

      const visitRules = (source: string, depth: number): void => {
        if (depth > MAX_CSS_DEPTH) return;
        if (rulesSeen >= MAX_CSS_RULES) return;

        let i = 0;
        let selectorStart = 0;
        let braceDepth = 0;
        let selector = "";
        let bodyStart = 0;
        let quote: string | null = null;

        while (i < source.length && rulesSeen < MAX_CSS_RULES) {
          const ch = source[i]!;
          if (quote) {
            if (ch === "\\" && i + 1 < source.length) {
              i += 2;
              continue;
            }
            if (ch === quote) quote = null;
            i += 1;
            continue;
          }
          if (ch === '"' || ch === "'") {
            quote = ch;
            i += 1;
            continue;
          }
          if (ch === "{") {
            if (braceDepth === 0) {
              selector = source.slice(selectorStart, i).trim();
              bodyStart = i + 1;
            }
            braceDepth += 1;
            i += 1;
            continue;
          }
          if (ch === "}") {
            if (braceDepth > 0) braceDepth -= 1;
            if (braceDepth === 0) {
              const body = source.slice(bodyStart, i);
              const sel = selector.trim();
              selector = "";
              selectorStart = i + 1;
              if (sel) {
                if (sel.startsWith("@")) {
                  visitRules(body, depth + 1);
                } else {
                  let spec = -1;
                  for (const part of sel.split(",")) {
                    const trimmed = part.trim();
                    if (!selectorMatchesRoot(trimmed)) continue;
                    spec = Math.max(spec, computeSpecificity(trimmed));
                  }
                  if (spec >= 0) {
                    rulesSeen += 1;
                    for (const rawDecl of body.split(";")) {
                      const decl = rawDecl.trim();
                      if (!decl) continue;
                      const colon = decl.indexOf(":");
                      if (colon === -1) continue;
                      const name = decl.slice(0, colon).trim();
                      const value = decl.slice(colon + 1).replace(/!important\b/i, "").trim();
                      if (!value) continue;
                      const lowerName = name.toLowerCase();

                      if (name.startsWith("--")) {
                        order += 1;
                        const next = { spec, order, value };
                        const prev = varCandidates.get(name) ?? null;
                        varCandidates.set(name, updateCandidate(prev, next));
                        continue;
                      }

                      if (lowerName === "width") {
                        order += 1;
                        bestWidth = updateCandidate(bestWidth, { spec, order, value });
                        continue;
                      }
                      if (lowerName === "height") {
                        order += 1;
                        bestHeight = updateCandidate(bestHeight, { spec, order, value });
                        continue;
                      }
                    }
                  }
                }
              }
            }
            i += 1;
            continue;
          }
          i += 1;
        }
      };

      visitRules(css, 0);
    };

    const decodeStyleBytes = (view: Uint8Array): string | null =>
      utf16Bom ? decodeUtf16(view, utf16LittleEndian) : decodeUtf8(view);

    const ingestStyleBytes = (view: Uint8Array): void => {
      const css = decodeStyleBytes(view);
      if (!css) return;
      ingestCss(css);
    };

    const MAX_STYLE_TAGS = 8;
    let styleTagsSeen = 0;

    if (utf16Bom) {
      const asciiLower = (cu: number): number => (cu >= 0x41 && cu <= 0x5a ? cu | 0x20 : cu);
      const isWhitespaceCu = (cu: number): boolean => cu === 0x20 || cu === 0x09 || cu === 0x0a || cu === 0x0d;

      const isStyleName = (start: number, end: number): boolean => {
        const nameLen = (end - start) / 2;
        if (nameLen === 5) {
          return (
            asciiLower(readUtf16CodeUnit(start)) === 0x73 &&
            asciiLower(readUtf16CodeUnit(start + 2)) === 0x74 &&
            asciiLower(readUtf16CodeUnit(start + 4)) === 0x79 &&
            asciiLower(readUtf16CodeUnit(start + 6)) === 0x6c &&
            asciiLower(readUtf16CodeUnit(start + 8)) === 0x65
          );
        }
        if (nameLen > 6 && readUtf16CodeUnit(end - 12) === 0x3a) {
          return (
            asciiLower(readUtf16CodeUnit(end - 10)) === 0x73 &&
            asciiLower(readUtf16CodeUnit(end - 8)) === 0x74 &&
            asciiLower(readUtf16CodeUnit(end - 6)) === 0x79 &&
            asciiLower(readUtf16CodeUnit(end - 4)) === 0x6c &&
            asciiLower(readUtf16CodeUnit(end - 2)) === 0x65
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

      const findStyleClose = (start: number): { contentEnd: number; afterEndTag: number } | null => {
        let i = start;
        while (i + 1 < len) {
          while (i + 1 < len && readUtf16CodeUnit(i) !== 0x3c) i += 2; // '<'
          if (i + 1 >= len || i + 3 >= len) return null;
          const next = readUtf16CodeUnit(i + 2);

          if (next === 0x21) {
            // Comment / CDATA / declaration.
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

            const end = findTagEnd(i);
            if (end === -1) return null;
            i = end;
            continue;
          }

          if (next === 0x3f) {
            // Processing instruction.
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
            const nameStart = i + 4;
            let nameEnd = nameStart;
            const maxNameEnd = Math.min(len - (len % 2), nameStart + 128);
            while (nameEnd + 1 < maxNameEnd) {
              const cu = readUtf16CodeUnit(nameEnd);
              if (isWhitespaceCu(cu) || cu === 0x2f || cu === 0x3e || cu === 0x3f) break;
              nameEnd += 2;
            }
            if (nameEnd > nameStart && isStyleName(nameStart, nameEnd)) {
              const end = findTagEnd(i);
              if (end === -1) return null;
              return { contentEnd: i, afterEndTag: end };
            }
            const end = findTagEnd(i);
            if (end === -1) return null;
            i = end;
            continue;
          }

          // Some other start tag inside style content; skip it.
          const end = findTagEnd(i);
          if (end === -1) return null;
          i = end;
        }
        return null;
      };

      let i = tagEnd;
      while (i + 1 < len && styleTagsSeen < MAX_STYLE_TAGS) {
        while (i + 1 < len && readUtf16CodeUnit(i) !== 0x3c) i += 2; // '<'
        if (i + 1 >= len || i + 3 >= len) break;

        const next = readUtf16CodeUnit(i + 2);
        if (next === 0x21 || next === 0x3f || next === 0x2f) {
          const end = findTagEnd(i);
          if (end === -1) break;
          i = end;
          continue;
        }

        const nameStart = i + 2;
        let nameEnd = nameStart;
        const maxNameEnd = Math.min(len - (len % 2), nameStart + 128);
        while (nameEnd + 1 < maxNameEnd) {
          const cu = readUtf16CodeUnit(nameEnd);
          if (isWhitespaceCu(cu) || cu === 0x2f || cu === 0x3e || cu === 0x3f) break;
          nameEnd += 2;
        }

        const startTagEnd = findTagEnd(i);
        if (startTagEnd === -1) break;

        if (nameEnd > nameStart && isStyleName(nameStart, nameEnd)) {
          const close = findStyleClose(startTagEnd);
          if (!close) {
            i = startTagEnd;
            continue;
          }
          ingestStyleBytes(bytes.subarray(startTagEnd, close.contentEnd));
          styleTagsSeen += 1;
          i = close.afterEndTag;
          continue;
        }

        i = startTagEnd;
      }
    } else {
      const asciiLower = (b: number): number => (b >= 0x41 && b <= 0x5a ? b | 0x20 : b);

      const isStyleName = (start: number, end: number): boolean => {
        const nameLen = end - start;
        if (nameLen === 5) {
          return (
            asciiLower(bytes[start]!) === 0x73 &&
            asciiLower(bytes[start + 1]!) === 0x74 &&
            asciiLower(bytes[start + 2]!) === 0x79 &&
            asciiLower(bytes[start + 3]!) === 0x6c &&
            asciiLower(bytes[start + 4]!) === 0x65
          );
        }
        if (nameLen > 6 && bytes[end - 6] === 0x3a) {
          return (
            asciiLower(bytes[end - 5]!) === 0x73 &&
            asciiLower(bytes[end - 4]!) === 0x74 &&
            asciiLower(bytes[end - 3]!) === 0x79 &&
            asciiLower(bytes[end - 2]!) === 0x6c &&
            asciiLower(bytes[end - 1]!) === 0x65
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

      const findStyleClose = (start: number): { contentEnd: number; afterEndTag: number } | null => {
        let i = start;
        while (i < len) {
          while (i < len && bytes[i] !== 0x3c) i += 1; // '<'
          if (i >= len || i + 1 >= len) return null;
          const next = bytes[i + 1]!;

          if (next === 0x21) {
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

            const end = findTagEnd(i);
            if (end === -1) return null;
            i = end;
            continue;
          }

          if (next === 0x3f) {
            // <? ... ?>
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
            const nameStart = i + 2;
            let nameEnd = nameStart;
            const maxNameEnd = Math.min(len, nameStart + 64);
            while (nameEnd < maxNameEnd) {
              const b = bytes[nameEnd]!;
              if (isWhitespaceByte(b) || b === 0x2f || b === 0x3e || b === 0x3f) break;
              nameEnd += 1;
            }
            if (nameEnd > nameStart && isStyleName(nameStart, nameEnd)) {
              const end = findTagEnd(i);
              if (end === -1) return null;
              return { contentEnd: i, afterEndTag: end };
            }
            const end = findTagEnd(i);
            if (end === -1) return null;
            i = end;
            continue;
          }

          const end = findTagEnd(i);
          if (end === -1) return null;
          i = end;
        }
        return null;
      };

      let i = tagEnd;
      while (i < len && styleTagsSeen < MAX_STYLE_TAGS) {
        while (i < len && bytes[i] !== 0x3c) i += 1; // '<'
        if (i >= len || i + 1 >= len) break;

        const next = bytes[i + 1]!;
        if (next === 0x21 || next === 0x3f || next === 0x2f) {
          const end = findTagEnd(i);
          if (end === -1) break;
          i = end;
          continue;
        }

        const nameStart = i + 1;
        let nameEnd = nameStart;
        const maxNameEnd = Math.min(len, nameStart + 64);
        while (nameEnd < maxNameEnd) {
          const b = bytes[nameEnd]!;
          if (isWhitespaceByte(b) || b === 0x2f || b === 0x3e || b === 0x3f) break;
          nameEnd += 1;
        }

        const startTagEnd = findTagEnd(i);
        if (startTagEnd === -1) break;

        if (nameEnd > nameStart && isStyleName(nameStart, nameEnd)) {
          const close = findStyleClose(startTagEnd);
          if (!close) {
            i = startTagEnd;
            continue;
          }
          ingestStyleBytes(bytes.subarray(startTagEnd, close.contentEnd));
          styleTagsSeen += 1;
          i = close.afterEndTag;
          continue;
        }

        i = startTagEnd;
      }
    }

    if ((width == null || height == null) && (bestWidth || bestHeight)) {
      const vars = new Map<string, string>();
      for (const [name, cand] of varCandidates) vars.set(name, cand.value);
      if (width == null && bestWidth) width = parseLength(resolveCssVars(bestWidth.value, vars));
      if (height == null && bestHeight) height = parseLength(resolveCssVars(bestHeight.value, vars));
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
