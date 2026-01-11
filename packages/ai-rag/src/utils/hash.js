const encoder = new TextEncoder();

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
  const bytes = encoder.encode(String(text));
  const subtle = globalThis.crypto?.subtle;
  if (subtle) {
    const digest = await subtle.digest("SHA-256", bytes);
    return bytesToHex(new Uint8Array(digest));
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
  let hash = 0xcbf29ce484222325n;
  const prime = 0x100000001b3n;
  for (const b of bytes) {
    hash ^= BigInt(b);
    hash = BigInt.asUintN(64, hash * prime);
  }
  return hash.toString(16).padStart(16, "0");
}
