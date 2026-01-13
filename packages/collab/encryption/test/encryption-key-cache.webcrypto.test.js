import assert from "node:assert/strict";
import test from "node:test";

test("index.ts (WebCrypto) encryption key cache is bounded (LRU) and clearable", async (t) => {
  // Force a tiny cache so we can deterministically test eviction semantics.
  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 2;

  const nodeCrypto = await import("node:crypto");

  // The browser/WebCrypto entrypoint requires `globalThis.crypto`. Some Node
  // versions (e.g. Node 18) don't expose it by default, so shim it from
  // `node:crypto`.webcrypto for the duration of this test.
  const hadGlobalCrypto = "crypto" in globalThis;
  if (!hadGlobalCrypto) {
    Object.defineProperty(globalThis, "crypto", {
      value: nodeCrypto.webcrypto,
      configurable: true,
      enumerable: true,
      writable: true,
    });
  }

  const mod = await import("../src/index.ts");
  const { encryptCellPlaintext, decryptCellPlaintext, clearEncryptionKeyCache } = mod;

  assert.equal(typeof encryptCellPlaintext, "function");
  assert.equal(typeof decryptCellPlaintext, "function");
  assert.equal(typeof clearEncryptionKeyCache, "function");

  const subtle = globalThis.crypto.subtle;
  assert.ok(subtle, "expected globalThis.crypto.subtle to exist for this test");

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
    if (!hadGlobalCrypto) {
      delete globalThis.crypto;
    }
  });

  /**
   * @param {string} keyId
   * @param {number} fill
   */
  const makeKey = (keyId, fill) => ({ keyId, keyBytes: new Uint8Array(32).fill(fill) });

  const context = { docId: "d1", sheetId: "Sheet1", row: 0, col: 0 };
  const plaintext = { value: "hello", formula: null };

  clearEncryptionKeyCache();
  importKeyCalls = 0;

  const key1 = makeKey("k1", 1);
  const encrypted = await encryptCellPlaintext({ plaintext, key: key1, context });
  const decrypted = await decryptCellPlaintext({ encrypted, key: key1, context });
  assert.deepEqual(decrypted, plaintext);
  assert.equal(importKeyCalls, 1, "expected decrypt to reuse cached CryptoKey");

  const key2 = makeKey("k2", 2);
  await encryptCellPlaintext({ plaintext, key: key2, context });
  assert.equal(importKeyCalls, 2);

  // Refresh key1 recency.
  await encryptCellPlaintext({ plaintext, key: key1, context });
  assert.equal(importKeyCalls, 2);

  // Add key3 -> should evict key2.
  const key3 = makeKey("k3", 3);
  await encryptCellPlaintext({ plaintext, key: key3, context });
  assert.equal(importKeyCalls, 3);

  // key2 should be evicted.
  const encrypted2 = await encryptCellPlaintext({ plaintext, key: key2, context });
  assert.equal(importKeyCalls, 4);
  const decrypted2 = await decryptCellPlaintext({ encrypted: encrypted2, key: key2, context });
  assert.deepEqual(decrypted2, plaintext);

  clearEncryptionKeyCache();
  const encrypted3 = await encryptCellPlaintext({ plaintext, key: key3, context });
  assert.equal(importKeyCalls, 5, "expected clearEncryptionKeyCache() to force re-import");
  const decrypted3 = await decryptCellPlaintext({ encrypted: encrypted3, key: key3, context });
  assert.deepEqual(decrypted3, plaintext);

  // If the max cache size is reduced at runtime, the accessed key should refresh
  // its recency *before* eviction is applied so it doesn't get evicted as the LRU
  // entry.
  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 3;
  clearEncryptionKeyCache();
  importKeyCalls = 0;

  await encryptCellPlaintext({ plaintext, key: key1, context });
  await encryptCellPlaintext({ plaintext, key: key2, context });
  await encryptCellPlaintext({ plaintext, key: key3, context });
  assert.equal(importKeyCalls, 3);

  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 2;
  await encryptCellPlaintext({ plaintext, key: key1, context });
  assert.equal(importKeyCalls, 3, "expected shrinking cache to not evict the accessed key");
  await encryptCellPlaintext({ plaintext, key: key2, context });
  assert.equal(importKeyCalls, 4, "expected shrinking cache to evict a different (LRU) key");

  // Disabling caching (max size = 0) should avoid retaining keys entirely.
  globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__ = 0;
  clearEncryptionKeyCache();
  importKeyCalls = 0;

  await encryptCellPlaintext({ plaintext, key: key1, context });
  await encryptCellPlaintext({ plaintext, key: key1, context });
  assert.equal(importKeyCalls, 2, "expected cache disabled to import key on every call");
});
