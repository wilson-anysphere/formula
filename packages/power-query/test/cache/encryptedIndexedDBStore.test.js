import assert from "node:assert/strict";
import { createCipheriv, createDecipheriv } from "node:crypto";
import test from "node:test";

import "fake-indexeddb/auto";

import { CacheManager } from "../../src/cache/cache.js";
import { EncryptedCacheStore } from "../../src/cache/encryptedStore.js";
import { IndexedDBCacheStore } from "../../src/cache/indexeddb.js";

/**
 * Deterministic AES-256-GCM provider for tests.
 *
 * @returns {import("../../src/cache/encryptedStore.js").CacheCryptoProvider}
 */
function createTestCryptoProvider() {
  const keyBytes = new Uint8Array(32).fill(8);
  let counter = 0;
  return {
    keyVersion: 1,
    async encryptBytes(plaintext, aad) {
      const iv = Buffer.alloc(12);
      iv.writeUInt32BE(counter++, 8);
      const cipher = createCipheriv("aes-256-gcm", Buffer.from(keyBytes), iv);
      if (aad) cipher.setAAD(Buffer.from(aad));
      const ciphertext = Buffer.concat([cipher.update(Buffer.from(plaintext)), cipher.final()]);
      const tag = cipher.getAuthTag();
      return { keyVersion: 1, iv: new Uint8Array(iv), tag: new Uint8Array(tag), ciphertext: new Uint8Array(ciphertext) };
    },
    async decryptBytes(payload, aad) {
      const decipher = createDecipheriv("aes-256-gcm", Buffer.from(keyBytes), Buffer.from(payload.iv));
      decipher.setAuthTag(Buffer.from(payload.tag));
      if (aad) decipher.setAAD(Buffer.from(aad));
      const plaintext = Buffer.concat([decipher.update(Buffer.from(payload.ciphertext)), decipher.final()]);
      return new Uint8Array(plaintext);
    },
  };
}

test("EncryptedCacheStore + IndexedDBCacheStore: stores ciphertext and roundtrips", async () => {
  const dbName = `pq-cache-encrypted-idb-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const underlying = new IndexedDBCacheStore({ dbName });
  const store = new EncryptedCacheStore({ store: underlying, crypto: createTestCryptoProvider(), storeId: "unit-test" });
  const cache = new CacheManager({ store, now: () => 0 });

  const secret = "idb-secret-substring";
  await cache.set("k1", { secret, bytes: new Uint8Array([1, 2, 3]) });

  const value = await cache.get("k1");
  assert.ok(value && typeof value === "object");
  assert.equal(value.secret, secret);
  assert.deepEqual(value.bytes, new Uint8Array([1, 2, 3]));

  const raw = await underlying.get("k1");
  assert.ok(raw);
  assert.ok(raw.value && typeof raw.value === "object");
  // @ts-ignore - test access
  assert.equal(raw.value.v, 2);
  // @ts-ignore - test access
  const ciphertext = raw.value.payload?.ciphertext;
  assert.ok(ciphertext instanceof Uint8Array || ciphertext instanceof ArrayBuffer);
  const ciphertextBytes = ciphertext instanceof Uint8Array ? ciphertext : new Uint8Array(ciphertext);
  assert.equal(Buffer.from(ciphertextBytes).includes(Buffer.from(secret)), false);

  // Close DB handles before deleting to keep fake-indexeddb happy.
  const db = await underlying.open();
  db.close();

  await new Promise((resolve, reject) => {
    const req = indexedDB.deleteDatabase(dbName);
    req.onsuccess = () => resolve(undefined);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB deleteDatabase failed"));
    req.onblocked = () => resolve(undefined);
  });
});
