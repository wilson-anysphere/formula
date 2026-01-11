const encoder = typeof TextEncoder !== "undefined" ? new TextEncoder() : null;

/**
 * @param {string} text
 */
function utf8Bytes(text) {
  const str = String(text);
  if (encoder) return encoder.encode(str);

  // Minimal UTF-8 encoder fallback for environments without TextEncoder.
  const encoded = encodeURIComponent(str);
  /** @type {number[]} */
  const bytes = [];
  for (let i = 0; i < encoded.length; i += 1) {
    const ch = encoded.charCodeAt(i);
    if (ch === 37 /* % */) {
      const hex = encoded.slice(i + 1, i + 3);
      bytes.push(Number.parseInt(hex, 16));
      i += 2;
      continue;
    }
    bytes.push(ch);
  }
  return Uint8Array.from(bytes);
}

/**
 * Fast, deterministic content hash intended for incremental indexing cache keys.
 *
 * Uses WebCrypto SHA-256 when available; falls back to FNV-1a 64-bit for
 * environments without WebCrypto.
 *
 * @param {string} text
 * @returns {Promise<string>} lowercase hex digest
 */
export async function contentHash(text) {
  const bytes = utf8Bytes(text);
  const subtle = globalThis.crypto?.subtle;
  if (subtle) {
    try {
      const digest = await subtle.digest("SHA-256", bytes);
      return bytesToHex(new Uint8Array(digest));
    } catch {
      // fall through to non-WebCrypto fallback
    }
  }

  // Extremely small fallback for environments without WebCrypto.
  return fnv1a64Hex(bytes);
}

/**
 * @param {string} text
 * @returns {Promise<string>} lowercase hex digest
 */
export async function sha256Hex(text) {
  return contentHash(text);
}

/**
 * @param {Uint8Array} bytes
 */
function bytesToHex(bytes) {
  let out = "";
  for (const b of bytes) out += b.toString(16).padStart(2, "0");
  return out;
}

/**
 * FNV-1a 64-bit, returned as 16-char lowercase hex.
 * @param {Uint8Array} bytes
 */
function fnv1a64Hex(bytes) {
  // Some JS runtimes support WebCrypto but not BigInt; provide a non-BigInt
  // fallback so `contentHash` still works anywhere.
  if (typeof BigInt === "undefined") {
    const h1 = fnv1a32(bytes, 0x811c9dc5);
    const h2 = fnv1a32(bytes, 0x811c9dc5 ^ 0x9e3779b9);
    return `${h1.toString(16).padStart(8, "0")}${h2.toString(16).padStart(8, "0")}`;
  }

  let hash = 0xcbf29ce484222325n;
  const prime = 0x100000001b3n;
  for (const b of bytes) {
    hash ^= BigInt(b);
    hash = BigInt.asUintN(64, hash * prime);
  }
  return hash.toString(16).padStart(16, "0");
}

/**
 * 32-bit FNV-1a.
 * @param {Uint8Array} bytes
 * @param {number} seed
 */
function fnv1a32(bytes, seed) {
  let hash = seed >>> 0;
  for (const b of bytes) {
    hash ^= b;
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}
