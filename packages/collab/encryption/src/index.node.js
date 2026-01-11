export const CELL_ENCRYPTION_VERSION = 1;
export const CELL_ENCRYPTION_ALG = "AES-256-GCM";

export const AES_256_KEY_BYTES = 32;
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;

let webCryptoPromise = null;
let webCryptoCached = null;

async function getWebCrypto() {
  if (webCryptoCached) return webCryptoCached;
  if (webCryptoPromise) return webCryptoPromise;

  const cryptoObj = globalThis.crypto;
  if (cryptoObj?.subtle && typeof cryptoObj.getRandomValues === "function") {
    webCryptoCached = cryptoObj;
    return webCryptoCached;
  }

  webCryptoPromise = (async () => {
    try {
      // Node.js: fall back to `crypto.webcrypto` on runtimes that don't expose
      // WebCrypto on `globalThis.crypto` (e.g. Node 18).
      //
      // This uses a dynamic import with a computed specifier so browser bundlers
      // don't try to resolve `node:crypto` for the desktop binder code path.
      const specifier = ["node", "crypto"].join(":");
      // eslint-disable-next-line no-undef
      const mod = await import(
        // eslint-disable-next-line no-undef
        /* @vite-ignore */ specifier
      );
      const webcrypto = mod?.webcrypto ?? mod?.default?.webcrypto ?? null;
      if (webcrypto?.subtle && typeof webcrypto.getRandomValues === "function") {
        webCryptoCached = webcrypto;
        return webCryptoCached;
      }
    } catch {
      // ignore and throw below
    }

    throw new Error(
      "WebCrypto is required for cell encryption (globalThis.crypto.subtle missing and Node crypto.webcrypto unavailable)"
    );
  })();

  return webCryptoPromise;
}

/**
 * @param {Uint8Array} keyBytes
 */
function assertKeyBytes(keyBytes) {
  if (!(keyBytes instanceof Uint8Array)) {
    throw new TypeError("keyBytes must be a Uint8Array");
  }
  if (keyBytes.byteLength !== AES_256_KEY_BYTES) {
    throw new RangeError(`keyBytes must be ${AES_256_KEY_BYTES} bytes (got ${keyBytes.byteLength})`);
  }
}

function isPlainObject(value) {
  return value != null && typeof value === "object" && value.constructor === Object;
}

function sortJson(value) {
  if (Array.isArray(value)) return value.map(sortJson);
  if (isPlainObject(value)) {
    const sorted = {};
    for (const key of Object.keys(value).sort()) {
      sorted[key] = sortJson(value[key]);
    }
    return sorted;
  }
  return value;
}

/**
 * Deterministic JSON encoding suitable for use as AAD / encryption context.
 *
 * This intentionally does *not* attempt to be a general-purpose canonicalization
 * algorithm. It exists so that `{ docId, sheetId, row, col }` produces identical
 * bytes across runtimes for AES-GCM additional authenticated data.
 *
 * @param {any} value
 */
export function canonicalJson(value) {
  return JSON.stringify(sortJson(value));
}

/**
 * @param {Uint8Array} bytes
 */
function bytesToBase64(bytes) {
  if (!(bytes instanceof Uint8Array)) {
    throw new TypeError("bytesToBase64 expects a Uint8Array");
  }
  // Node: Buffer path.
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");

  let bin = "";
  for (const b of bytes) bin += String.fromCharCode(b);
  // eslint-disable-next-line no-undef
  return btoa(bin);
}

/**
 * @param {string} value
 */
function base64ToBytes(value) {
  if (typeof value !== "string") {
    throw new TypeError("base64ToBytes expects a base64 string");
  }
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(value, "base64"));

  // eslint-disable-next-line no-undef
  const bin = atob(value);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
  return out;
}

/**
 * @param {string} text
 */
function utf8Encode(text) {
  if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(text);
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(String(text), "utf8"));
  throw new Error("No UTF-8 encoder available (TextEncoder missing)");
}

/**
 * @param {Uint8Array} bytes
 */
function utf8Decode(bytes) {
  if (typeof TextDecoder !== "undefined") return new TextDecoder().decode(bytes);
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("utf8");
  throw new Error("No UTF-8 decoder available (TextDecoder missing)");
}

function concatBytes(a, b) {
  const out = new Uint8Array(a.byteLength + b.byteLength);
  out.set(a, 0);
  out.set(b, a.byteLength);
  return out;
}

/**
 * @param {{ docId: string, sheetId: string, row: number, col: number }} context
 */
function aadBytesFromContext(context) {
  return utf8Encode(canonicalJson(context));
}

/**
 * @param {unknown} value
 */
export function isEncryptedCellPayload(value) {
  if (!value || typeof value !== "object") return false;
  const v = value;
  return (
    v.v === CELL_ENCRYPTION_VERSION &&
    v.alg === CELL_ENCRYPTION_ALG &&
    typeof v.keyId === "string" &&
    typeof v.ivBase64 === "string" &&
    typeof v.tagBase64 === "string" &&
    typeof v.ciphertextBase64 === "string"
  );
}

const keyCache = new Map();

async function importAesGcmKey(key) {
  const cached = keyCache.get(key.keyId);
  if (cached) return cached;

  assertKeyBytes(key.keyBytes);
  const subtle = (await getWebCrypto()).subtle;
  const cryptoKey = await subtle.importKey(
    "raw",
    key.keyBytes,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"]
  );
  keyCache.set(key.keyId, cryptoKey);
  return cryptoKey;
}

/**
 * @param {{
 *   plaintext: { value: any, formula: string | null, format?: any },
 *   key: { keyId: string, keyBytes: Uint8Array },
 *   context: { docId: string, sheetId: string, row: number, col: number },
 * }} opts
 */
export async function encryptCellPlaintext(opts) {
  const { plaintext, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  const cryptoObj = await getWebCrypto();
  const iv = cryptoObj.getRandomValues(new Uint8Array(AES_GCM_IV_BYTES));
  const aad = aadBytesFromContext(context);
  const bytes = utf8Encode(JSON.stringify(plaintext));

  const cryptoKey = await importAesGcmKey(key);
  const ciphertextWithTag = new Uint8Array(
    await cryptoObj.subtle.encrypt(
      { name: "AES-GCM", iv, additionalData: aad, tagLength: AES_GCM_TAG_BYTES * 8 },
      cryptoKey,
      bytes
    )
  );

  const tag = ciphertextWithTag.slice(ciphertextWithTag.byteLength - AES_GCM_TAG_BYTES);
  const ciphertext = ciphertextWithTag.slice(0, ciphertextWithTag.byteLength - AES_GCM_TAG_BYTES);

  return {
    v: CELL_ENCRYPTION_VERSION,
    alg: CELL_ENCRYPTION_ALG,
    keyId: key.keyId,
    ivBase64: bytesToBase64(iv),
    tagBase64: bytesToBase64(tag),
    ciphertextBase64: bytesToBase64(ciphertext),
  };
}

/**
 * @param {{
 *   encrypted: { v: number, alg: string, keyId: string, ivBase64: string, tagBase64: string, ciphertextBase64: string },
 *   key: { keyId: string, keyBytes: Uint8Array },
 *   context: { docId: string, sheetId: string, row: number, col: number },
 * }} opts
 */
export async function decryptCellPlaintext(opts) {
  const { encrypted, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  if (encrypted.keyId !== key.keyId) {
    throw new Error(`Key id mismatch (payload=${encrypted.keyId}, resolver=${key.keyId})`);
  }

  const cryptoObj = await getWebCrypto();

  const iv = base64ToBytes(encrypted.ivBase64);
  const tag = base64ToBytes(encrypted.tagBase64);
  const ciphertext = base64ToBytes(encrypted.ciphertextBase64);
  const aad = aadBytesFromContext(context);

  const cryptoKey = await importAesGcmKey(key);
  const combined = concatBytes(ciphertext, tag);

  const plaintextBytes = new Uint8Array(
    await cryptoObj.subtle.decrypt(
      { name: "AES-GCM", iv, additionalData: aad, tagLength: AES_GCM_TAG_BYTES * 8 },
      cryptoKey,
      combined
    )
  );

  return JSON.parse(utf8Decode(plaintextBytes));
}
