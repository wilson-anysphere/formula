import assert from "node:assert/strict";
import { createCipheriv, createDecipheriv } from "node:crypto";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { EncryptedCacheStore } from "../../src/cache/encryptedStore.js";
import { FileSystemCacheStore } from "../../src/cache/filesystem.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";

/**
 * @param {{ keyBytes?: Uint8Array, keyVersion?: number }} [options]
 * @returns {import("../../src/cache/encryptedStore.js").CacheCryptoProvider}
 */
function createTestCryptoProvider(options = {}) {
  const keyBytes = options.keyBytes ?? new Uint8Array(32).fill(7);
  const keyVersion = options.keyVersion ?? 1;
  let counter = 0;

  return {
    keyVersion,
    async encryptBytes(plaintext, aad) {
      const iv = Buffer.alloc(12);
      iv.writeUInt32BE(counter++, 8);
      const cipher = createCipheriv("aes-256-gcm", Buffer.from(keyBytes), iv);
      if (aad) cipher.setAAD(Buffer.from(aad));
      const ciphertext = Buffer.concat([cipher.update(Buffer.from(plaintext)), cipher.final()]);
      const tag = cipher.getAuthTag();
      return { keyVersion, iv: new Uint8Array(iv), tag: new Uint8Array(tag), ciphertext: new Uint8Array(ciphertext) };
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

test("EncryptedCacheStore: roundtrips structured values and does not store plaintext", async () => {
  const underlying = new MemoryCacheStore();
  const crypto = createTestCryptoProvider();
  const store = new EncryptedCacheStore({ store: underlying, crypto, storeId: "unit-test" });
  const cache = new CacheManager({ store, now: () => 0 });

  const secret = "my-secret-substring";
  const value = { secret, bytes: new Uint8Array([1, 2, 3]), nested: { ok: true } };

  await cache.set("k1", value);

  const roundtrip = await cache.get("k1");
  assert.ok(roundtrip && typeof roundtrip === "object");
  assert.equal(roundtrip.secret, secret);
  assert.ok(roundtrip.bytes instanceof Uint8Array);
  assert.deepEqual(roundtrip.bytes, new Uint8Array([1, 2, 3]));

  const stored = underlying.map.get("k1");
  assert.ok(stored);
  // Ensure ciphertext doesn't trivially contain plaintext substrings.
  assert.ok(stored.value && typeof stored.value === "object");
  // @ts-ignore - test access
  const ciphertext = stored.value.payload?.ciphertext;
  assert.ok(ciphertext instanceof Uint8Array);
  assert.equal(Buffer.from(ciphertext).includes(Buffer.from(secret)), false);
});

test("EncryptedCacheStore: corrupt ciphertext is treated as a miss and deleted", async () => {
  const underlying = new MemoryCacheStore();
  const crypto = createTestCryptoProvider();
  const store = new EncryptedCacheStore({ store: underlying, crypto, storeId: "unit-test" });

  await store.set("k1", { value: { a: 1 }, createdAtMs: 0, expiresAtMs: null });

  const stored = underlying.map.get("k1");
  assert.ok(stored);
  // @ts-ignore - test access
  stored.value.payload.ciphertext[0] ^= 0xff;

  assert.equal(await store.get("k1"), null);
  assert.equal(underlying.map.has("k1"), false);
});

test("EncryptedCacheStore: wrong key is treated as a miss and deleted", async () => {
  const underlying = new MemoryCacheStore();
  const cryptoA = createTestCryptoProvider({ keyBytes: new Uint8Array(32).fill(1) });
  const cryptoB = createTestCryptoProvider({ keyBytes: new Uint8Array(32).fill(2) });

  const storeA = new EncryptedCacheStore({ store: underlying, crypto: cryptoA, storeId: "unit-test" });
  await storeA.set("k1", { value: { a: 1 }, createdAtMs: 0, expiresAtMs: null });

  const storeB = new EncryptedCacheStore({ store: underlying, crypto: cryptoB, storeId: "unit-test" });
  assert.equal(await storeB.get("k1"), null);
  assert.equal(underlying.map.has("k1"), false);
});

test("EncryptedCacheStore: storeId mismatch (AAD) is treated as a miss and deleted", async () => {
  const underlying = new MemoryCacheStore();
  const crypto = createTestCryptoProvider({ keyBytes: new Uint8Array(32).fill(3) });

  const storeA = new EncryptedCacheStore({ store: underlying, crypto, storeId: "store-a" });
  await storeA.set("k1", { value: { secret: "x" }, createdAtMs: 0, expiresAtMs: null });

  const storeB = new EncryptedCacheStore({ store: underlying, crypto, storeId: "store-b" });
  assert.equal(await storeB.get("k1"), null);
  assert.equal(underlying.map.has("k1"), false);
});

test("EncryptedCacheStore: plaintext entries are treated as misses", async () => {
  const underlying = new MemoryCacheStore();
  await underlying.set("k1", { value: { secret: "plaintext" }, createdAtMs: 0, expiresAtMs: null });

  const store = new EncryptedCacheStore({ store: underlying, crypto: createTestCryptoProvider(), storeId: "unit-test" });
  assert.equal(await store.get("k1"), null);
  assert.equal(underlying.map.has("k1"), false);
});

test("EncryptedCacheStore: unknown envelope versions are treated as misses but retained", async () => {
  const underlying = new MemoryCacheStore();
  await underlying.set("k1", {
    value: { __pq_cache_encrypted: "power-query-cache-encrypted", v: 999, payload: { future: true } },
    createdAtMs: 0,
    expiresAtMs: null,
  });

  const store = new EncryptedCacheStore({ store: underlying, crypto: createTestCryptoProvider(), storeId: "unit-test" });
  assert.equal(await store.get("k1"), null);
  assert.equal(underlying.map.has("k1"), true);
});

test("EncryptedCacheStore + FileSystemCacheStore: persists encrypted blobs (no plaintext)", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-encrypted-"));

  try {
    const underlying = new FileSystemCacheStore({ directory: cacheDir });
    const store = new EncryptedCacheStore({ store: underlying, crypto: createTestCryptoProvider(), storeId: "unit-test" });
    const cache = new CacheManager({ store, now: () => 0 });

    const secret = "disk-secret-substring-0123456789";
    await cache.set("k1", { secret, bytes: new Uint8Array([4, 5, 6]) });

    const roundtrip = await cache.get("k1");
    assert.ok(roundtrip && typeof roundtrip === "object");
    assert.equal(roundtrip.secret, secret);
    assert.deepEqual(roundtrip.bytes, new Uint8Array([4, 5, 6]));

    const files = await readdir(cacheDir);
    assert.ok(files.some((name) => name.endsWith(".bin")), "encrypted filesystem cache should create a .bin blob");

    const { jsonPath, binPath } = await underlying.pathsForKey("k1");
    const jsonText = await readFile(jsonPath, "utf8");
    assert.equal(jsonText.includes(secret), false);

    const binBytes = await readFile(binPath);
    assert.equal(binBytes.includes(Buffer.from(secret)), false);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

test("CacheManager.pruneExpired forwards to the store and swallows errors", async () => {
  let calledWith = null;
  /** @type {import("../../src/cache/cache.js").CacheStore} */
  const store = {
    get: async () => null,
    set: async () => {},
    delete: async () => {},
    pruneExpired: async (nowMs) => {
      calledWith = nowMs ?? null;
      throw new Error("boom");
    },
  };

  const cache = new CacheManager({ store, now: () => 123 });
  await cache.pruneExpired();
  assert.equal(calledWith, 123);
});
