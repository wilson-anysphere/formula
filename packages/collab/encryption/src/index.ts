import {
  AES_GCM_IV_BYTES,
  AES_GCM_TAG_BYTES,
  CELL_ENCRYPTION_ALG,
  CELL_ENCRYPTION_VERSION,
  aadBytesFromContext,
  assertKeyBytes,
  assertTagBytes,
  assertIvBytes,
  base64ToBytes,
  bytesToBase64,
  concatBytes,
  utf8Decode,
  utf8Encode,
  type CellEncryptionContext,
  type CellEncryptionKey,
  type CellPlaintext,
  type EncryptedCellPayloadV1,
} from "./shared.ts";

function requireCrypto(): Crypto {
  const cryptoObj = globalThis.crypto;
  if (!cryptoObj?.getRandomValues) {
    throw new Error("WebCrypto is required for cell encryption (globalThis.crypto.getRandomValues missing)");
  }
  return cryptoObj;
}

function requireSubtleCrypto(): SubtleCrypto {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new Error("WebCrypto SubtleCrypto is required for cell encryption (globalThis.crypto.subtle missing)");
  }
  return subtle;
}

// Cache imported `CryptoKey`s by keyId so repeated encrypt/decrypt operations don't
// pay the cost of `subtle.importKey` each time.
//
// IMPORTANT: Imported CryptoKeys can retain sensitive key material. Keep the cache
// bounded so it cannot grow without limit (e.g. in enterprise deployments with
// many documents / keys).
//
// Configuration:
// - `globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__` (number)
const DEFAULT_ENCRYPTION_KEY_CACHE_MAX_SIZE = 256;
const keyCache = new Map<string, CryptoKey>();

function normalizeCacheMaxSize(value: unknown): number | null {
  const num = typeof value === "string" ? Number.parseInt(value, 10) : Number(value);
  if (!Number.isFinite(num) || num < 0) return null;
  return Math.trunc(num);
}

function getEncryptionKeyCacheMaxSize(): number {
  const fromGlobal = normalizeCacheMaxSize((globalThis as any)?.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__);
  if (fromGlobal != null) return fromGlobal;
  return DEFAULT_ENCRYPTION_KEY_CACHE_MAX_SIZE;
}

function touchEncryptionKeyCacheEntry(keyId: string, cryptoKey: CryptoKey): void {
  // Refresh LRU order (Map iterates in insertion order).
  keyCache.delete(keyId);
  keyCache.set(keyId, cryptoKey);
}

function enforceEncryptionKeyCacheLimit(maxSize: number = getEncryptionKeyCacheMaxSize()): void {
  // Allow explicitly disabling caching by setting max size to 0.
  if (maxSize === 0) {
    keyCache.clear();
    return;
  }

  while (keyCache.size > maxSize) {
    const oldestKeyId = keyCache.keys().next().value;
    if (oldestKeyId === undefined) return;
    keyCache.delete(oldestKeyId);
  }
}

export function clearEncryptionKeyCache(): void {
  keyCache.clear();
}

async function importAesGcmKey(key: CellEncryptionKey): Promise<CryptoKey> {
  const maxSize = getEncryptionKeyCacheMaxSize();

  // Allow explicitly disabling caching by setting max size to 0.
  // Clear any previously cached keys so key material is not retained.
  if (maxSize === 0) {
    keyCache.clear();
    assertKeyBytes(key.keyBytes);
    const subtle = requireSubtleCrypto();
    return await subtle.importKey("raw", key.keyBytes, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  }

  const cached = keyCache.get(key.keyId);
  if (cached) {
    touchEncryptionKeyCacheEntry(key.keyId, cached);
    // If max size decreased at runtime, evict after refreshing recency so the
    // accessed key is not treated as least-recently-used.
    enforceEncryptionKeyCacheLimit(maxSize);
    return cached;
  }

  assertKeyBytes(key.keyBytes);
  const subtle = requireSubtleCrypto();
  const cryptoKey = await subtle.importKey("raw", key.keyBytes, { name: "AES-GCM", length: 256 }, false, [
    "encrypt",
    "decrypt",
  ]);

  touchEncryptionKeyCacheEntry(key.keyId, cryptoKey);
  enforceEncryptionKeyCacheLimit(maxSize);
  return cryptoKey;
}

export async function encryptCellPlaintext(opts: {
  plaintext: CellPlaintext;
  key: CellEncryptionKey;
  context: CellEncryptionContext;
}): Promise<EncryptedCellPayloadV1> {
  const { plaintext, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  const cryptoObj = requireCrypto();
  const iv = cryptoObj.getRandomValues(new Uint8Array(AES_GCM_IV_BYTES));
  assertIvBytes(iv);
  const aad = aadBytesFromContext(context);

  const bytes = utf8Encode(JSON.stringify(plaintext));

  const cryptoKey = await importAesGcmKey(key);
  const subtle = requireSubtleCrypto();
  const ciphertextWithTag = new Uint8Array(
    await subtle.encrypt(
      { name: "AES-GCM", iv, additionalData: aad, tagLength: AES_GCM_TAG_BYTES * 8 },
      cryptoKey,
      bytes
    )
  );

  const tag = ciphertextWithTag.slice(ciphertextWithTag.byteLength - AES_GCM_TAG_BYTES);
  const ciphertext = ciphertextWithTag.slice(0, ciphertextWithTag.byteLength - AES_GCM_TAG_BYTES);

  assertTagBytes(tag);

  return {
    v: CELL_ENCRYPTION_VERSION,
    alg: CELL_ENCRYPTION_ALG,
    keyId: key.keyId,
    ivBase64: bytesToBase64(iv),
    tagBase64: bytesToBase64(tag),
    ciphertextBase64: bytesToBase64(ciphertext),
  };
}

export async function decryptCellPlaintext(opts: {
  encrypted: EncryptedCellPayloadV1;
  key: CellEncryptionKey;
  context: CellEncryptionContext;
}): Promise<CellPlaintext> {
  const { encrypted, key, context } = opts;
  assertKeyBytes(key.keyBytes);

  if (encrypted.keyId !== key.keyId) {
    throw new Error(`Key id mismatch (payload=${encrypted.keyId}, resolver=${key.keyId})`);
  }

  const iv = base64ToBytes(encrypted.ivBase64);
  const tag = base64ToBytes(encrypted.tagBase64);
  const ciphertext = base64ToBytes(encrypted.ciphertextBase64);
  assertIvBytes(iv);
  assertTagBytes(tag);

  const aad = aadBytesFromContext(context);

  const cryptoKey = await importAesGcmKey(key);
  const subtle = requireSubtleCrypto();
  const combined = concatBytes(ciphertext, tag);
  const plaintextBytes = new Uint8Array(
    await subtle.decrypt(
      { name: "AES-GCM", iv, additionalData: aad, tagLength: AES_GCM_TAG_BYTES * 8 },
      cryptoKey,
      combined
    )
  );

  const parsed = JSON.parse(utf8Decode(plaintextBytes));
  return parsed as CellPlaintext;
}

export {
  AES_256_KEY_BYTES,
  AES_GCM_IV_BYTES,
  AES_GCM_TAG_BYTES,
  CELL_ENCRYPTION_ALG,
  CELL_ENCRYPTION_VERSION,
  aadBytesFromContext,
  base64ToBytes,
  bytesToBase64,
  concatBytes,
  canonicalJson,
  isEncryptedCellPayload,
  utf8Decode,
  utf8Encode,
} from "./shared.ts";
export type {
  CellEncryptionContext,
  CellEncryptionKey,
  CellPlaintext,
  EncryptedCellPayload,
  EncryptedCellPayloadV1,
} from "./shared.ts";
