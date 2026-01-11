import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { EncryptedCacheStore } from "../../src/cache/encryptedStore.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { createNodeCryptoCacheProvider } from "../../src/cache/nodeCryptoProvider.js";

test("createNodeCryptoCacheProvider: roundtrips via EncryptedCacheStore", async () => {
  const cryptoProvider = createNodeCryptoCacheProvider({ keyVersion: 1, keyBytes: new Uint8Array(32).fill(4) });

  const underlying = new MemoryCacheStore();
  const store = new EncryptedCacheStore({ store: underlying, crypto: cryptoProvider, storeId: "unit-test" });
  const cache = new CacheManager({ store, now: () => 0 });

  const secret = "nodecrypto-secret-substring";
  await cache.set("k1", { secret, bytes: new Uint8Array([1, 2, 3]) });
  const value = await cache.get("k1");
  assert.ok(value && typeof value === "object");
  assert.equal(value.secret, secret);
  assert.deepEqual(value.bytes, new Uint8Array([1, 2, 3]));

  const stored = underlying.map.get("k1");
  assert.ok(stored);
  // @ts-ignore - test access
  const ciphertext = stored.value.payload?.ciphertext;
  assert.ok(ciphertext instanceof Uint8Array);
  assert.equal(Buffer.from(ciphertext).includes(Buffer.from(secret)), false);
});

