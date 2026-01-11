import crypto from "node:crypto";

export const CELL_ENCRYPTION_VERSION = 1;
export const CELL_ENCRYPTION_ALG = "AES-256-GCM";

export const AES_256_KEY_BYTES = 32;
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;

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
 * @param {{ docId: string, sheetId: string, row: number, col: number }} context
 */
function aadFromContext(context) {
  return Buffer.from(canonicalJson(context), "utf8");
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

  const iv = crypto.randomBytes(AES_GCM_IV_BYTES);
  const aad = aadFromContext(context);

  const bytes = Buffer.from(JSON.stringify(plaintext), "utf8");

  const cipher = crypto.createCipheriv("aes-256-gcm", Buffer.from(key.keyBytes), iv, {
    authTagLength: AES_GCM_TAG_BYTES,
  });
  cipher.setAAD(aad);
  const ciphertext = Buffer.concat([cipher.update(bytes), cipher.final()]);
  const tag = cipher.getAuthTag();

  return {
    v: CELL_ENCRYPTION_VERSION,
    alg: CELL_ENCRYPTION_ALG,
    keyId: key.keyId,
    ivBase64: iv.toString("base64"),
    tagBase64: tag.toString("base64"),
    ciphertextBase64: ciphertext.toString("base64"),
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

  const iv = Buffer.from(encrypted.ivBase64, "base64");
  const tag = Buffer.from(encrypted.tagBase64, "base64");
  const ciphertext = Buffer.from(encrypted.ciphertextBase64, "base64");
  const aad = aadFromContext(context);

  const decipher = crypto.createDecipheriv("aes-256-gcm", Buffer.from(key.keyBytes), iv, {
    authTagLength: AES_GCM_TAG_BYTES,
  });
  decipher.setAuthTag(tag);
  decipher.setAAD(aad);
  const plaintextBytes = Buffer.concat([decipher.update(ciphertext), decipher.final()]);

  return JSON.parse(plaintextBytes.toString("utf8"));
}

