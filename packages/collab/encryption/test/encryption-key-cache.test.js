import assert from "node:assert/strict";
import test from "node:test";

test("index.node.js encryption key cache is bounded (LRU) and clearable", async (t) => {
  // Keep the cache tiny so we can deterministically test eviction semantics without
  // importing hundreds of keys.
  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 2;

  const mod = await import("../src/index.node.js");
  const { encryptCellPlaintext, decryptCellPlaintext, clearEncryptionKeyCache } = mod;

  assert.equal(typeof encryptCellPlaintext, "function");
  assert.equal(typeof decryptCellPlaintext, "function");
  assert.equal(typeof clearEncryptionKeyCache, "function");

  // Node 18 may not expose WebCrypto on `globalThis.crypto`. The implementation
  // under test falls back to `node:crypto`'s `webcrypto` in that case, so patch
  // the `SubtleCrypto#importKey` method on whichever WebCrypto instance the
  // module will use.
  const nodeCrypto = await import("node:crypto");
  const globalCrypto = globalThis.crypto;
  const cryptoObj =
    globalCrypto?.subtle && typeof globalCrypto.getRandomValues === "function" ? globalCrypto : nodeCrypto.webcrypto;
  assert.ok(cryptoObj?.subtle, "expected WebCrypto (crypto.subtle) to be available for this test");

  const subtle = cryptoObj.subtle;
  const originalImportKey = subtle.importKey;
  let importKeyCalls = 0;
  subtle.importKey = async (...args) => {
    importKeyCalls += 1;
    return await originalImportKey.apply(subtle, args);
  };

  t.after(() => {
    subtle.importKey = originalImportKey;
    delete globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__;
    clearEncryptionKeyCache();
  });

  /**
   * @param {string} keyId
   * @param {number} fill
   */
  const makeKey = (keyId, fill) => ({ keyId, keyBytes: new Uint8Array(32).fill(fill) });

  const context = { docId: "d1", sheetId: "Sheet1", row: 0, col: 0 };
  const plaintext = { value: "hello", formula: null };

  // Start from a clean slate.
  clearEncryptionKeyCache();
  importKeyCalls = 0;

  // Basic correctness + cache hit: encrypt/decrypt should only import once.
  const key1 = makeKey("k1", 1);
  const encrypted = await encryptCellPlaintext({ plaintext, key: key1, context });
  const decrypted = await decryptCellPlaintext({ encrypted, key: key1, context });
  assert.deepEqual(decrypted, plaintext);
  assert.equal(importKeyCalls, 1, "expected decrypt to reuse cached CryptoKey");

  // Populate cache to capacity.
  const key2 = makeKey("k2", 2);
  await encryptCellPlaintext({ plaintext, key: key2, context });
  assert.equal(importKeyCalls, 2);

  // Access key1 to make it most-recently-used.
  await encryptCellPlaintext({ plaintext, key: key1, context });
  assert.equal(importKeyCalls, 2, "expected cache hit for key1");

  // Adding key3 should evict key2 (least-recently-used).
  const key3 = makeKey("k3", 3);
  await encryptCellPlaintext({ plaintext, key: key3, context });
  assert.equal(importKeyCalls, 3);

  // key2 should have been evicted, so using it again should trigger a re-import.
  await encryptCellPlaintext({ plaintext, key: key2, context });
  assert.equal(importKeyCalls, 4, "expected key2 to be evicted once over capacity");

  // Clearing should drop all cached keys.
  clearEncryptionKeyCache();
  await encryptCellPlaintext({ plaintext, key: key3, context });
  assert.equal(importKeyCalls, 5, "expected clearEncryptionKeyCache() to force re-import");

  // Disabling caching (max size = 0) should avoid retaining keys entirely.
  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 0;
  clearEncryptionKeyCache();
  importKeyCalls = 0;

  await encryptCellPlaintext({ plaintext, key: key1, context });
  await encryptCellPlaintext({ plaintext, key: key1, context });
  assert.equal(importKeyCalls, 2, "expected cache disabled to import key on every call");
});
