/**
 * OAuth2 PKCE helpers.
 *
 * These utilities avoid Node-only dependencies so they can run in browsers,
 * web workers, and Node.
 *
 * @see RFC 7636 - Proof Key for Code Exchange by OAuth Public Clients
 */

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function base64UrlEncode(bytes) {
  /** @type {string} */
  let base64;
  // Node.js
  if (typeof Buffer !== "undefined") {
    base64 = Buffer.from(bytes).toString("base64");
  } else {
    // Browser
    let binary = "";
    for (const b of bytes) binary += String.fromCharCode(b);
    // eslint-disable-next-line no-undef
    base64 = btoa(binary);
  }
  return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

/**
 * @param {number} length
 * @returns {Promise<Uint8Array>}
 */
async function randomBytes(length) {
  const cryptoObj = /** @type {any} */ (globalThis.crypto);
  if (cryptoObj?.getRandomValues) {
    const out = new Uint8Array(length);
    cryptoObj.getRandomValues(out);
    return out;
  }

  const nodeCrypto = await import("node:crypto");
  return new Uint8Array(nodeCrypto.randomBytes(length));
}

/**
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function sha256(bytes) {
  const cryptoObj = /** @type {any} */ (globalThis.crypto);
  if (cryptoObj?.subtle?.digest) {
    const digest = await cryptoObj.subtle.digest("SHA-256", bytes);
    return new Uint8Array(digest);
  }

  const nodeCrypto = await import("node:crypto");
  const hash = nodeCrypto.createHash("sha256").update(Buffer.from(bytes)).digest();
  return new Uint8Array(hash);
}

/**
 * Create a PKCE code verifier.
 *
 * The returned string is URL-safe and suitable for use as the `code_verifier`
 * parameter when exchanging an authorization code.
 *
 * @param {{ byteLength?: number } | undefined} [options]
 * @returns {Promise<string>}
 */
export async function createCodeVerifier(options = {}) {
  // 32 bytes => 43 chars when base64url encoded (no padding), which meets the
  // RFC 7636 minimum verifier length.
  const byteLength = options.byteLength ?? 32;
  const bytes = await randomBytes(byteLength);
  return base64UrlEncode(bytes);
}

/**
 * Create a PKCE code challenge (S256).
 *
 * @param {string} verifier
 * @returns {Promise<string>}
 */
export async function createCodeChallenge(verifier) {
  const bytes = new TextEncoder().encode(verifier);
  const digest = await sha256(bytes);
  return base64UrlEncode(digest);
}

