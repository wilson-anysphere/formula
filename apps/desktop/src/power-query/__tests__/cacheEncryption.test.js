import assert from "node:assert/strict";
import { Buffer } from "node:buffer";
import test from "node:test";

import { createDesktopQueryEngine } from "../engine.ts";
import { EncryptedCacheStore } from "@formula/power-query";

let indexedDbAvailable = true;
/** @type {import("fake-indexeddb").indexedDB | null} */
let fakeIndexedDB = null;
/** @type {import("fake-indexeddb").IDBKeyRange | null} */
let fakeIDBKeyRange = null;
try {
  const mod = await import("fake-indexeddb");
  // @ts-ignore - `fake-indexeddb` has a default export map; `indexedDB` is the common entry.
  fakeIndexedDB = mod.indexedDB ?? null;
  // @ts-ignore - runtime export
  fakeIDBKeyRange = mod.IDBKeyRange ?? null;
  indexedDbAvailable = Boolean(fakeIndexedDB);
} catch {
  indexedDbAvailable = false;
}

const webCryptoAvailable = Boolean(globalThis.crypto?.subtle && globalThis.crypto?.getRandomValues);

const DB_NAME = "formula-power-query-cache-encrypted-v1";

/**
 * Delete an IndexedDB database by name (best-effort).
 * @param {string} name
 */
async function deleteDatabase(name) {
  if (!globalThis.indexedDB) return;
  await new Promise((resolve, reject) => {
    const req = globalThis.indexedDB.deleteDatabase(name);
    req.onsuccess = () => resolve(undefined);
    req.onerror = () => reject(req.error ?? new Error(`IndexedDB deleteDatabase failed for ${name}`));
    req.onblocked = () => resolve(undefined);
  });
}

test(
  "createDesktopQueryEngine encrypts Power Query IndexedDB cache entries at rest",
  { skip: !indexedDbAvailable || !webCryptoAvailable },
  async () => {
    const originalTauri = globalThis.__TAURI__;
    const originalIndexedDB = globalThis.indexedDB;
    // Some environments also expose IDBKeyRange; restore it for hygiene.
    const originalIDBKeyRange = globalThis.IDBKeyRange;

    /** @type {{ cmd: string, args: any }[]} */
    const calls = [];

    // Ensure a clean starting point.
    globalThis.indexedDB = fakeIndexedDB;
    if (fakeIDBKeyRange) {
      globalThis.IDBKeyRange = fakeIDBKeyRange;
    }
    await deleteDatabase(DB_NAME).catch(() => {});

    // 32-byte deterministic key for tests.
    const keyBytes = Buffer.alloc(32, 7);
    const keyBase64 = keyBytes.toString("base64");

    globalThis.__TAURI__ = {
      core: {
        invoke: async (cmd, args) => {
          calls.push({ cmd, args });
          if (cmd === "power_query_cache_key_get_or_create") {
            return { keyVersion: 1, keyBase64 };
          }
          throw new Error(`Unexpected invoke: ${cmd}`);
        },
      },
    };

    try {
      const engine = createDesktopQueryEngine();
      assert.ok(engine && typeof engine === "object");
      // @ts-ignore - runtime access
      const cache = engine.cache;
      assert.ok(cache, "expected desktop query engine to have a CacheManager");

      const secret = `pq-cache-secret-${Date.now()}-${Math.random().toString(16).slice(2)}`;
      await cache.set("k1", { secret, bytes: new Uint8Array([1, 2, 3]) });
      const roundtripped = await cache.get("k1");
      assert.ok(roundtripped && typeof roundtripped === "object");
      // @ts-ignore - runtime access
      assert.equal(roundtripped.secret, secret);
      // @ts-ignore - runtime access
      assert.deepEqual(roundtripped.bytes, new Uint8Array([1, 2, 3]));

      // @ts-ignore - runtime access
      const store = cache.store;
      assert.ok(store instanceof EncryptedCacheStore);
      const raw = await store.store.get("k1");
      assert.ok(raw);

      // @ts-ignore - envelope access
      assert.equal(raw.value?.v, 2);
      // @ts-ignore - envelope access
      const ciphertext = raw.value?.payload?.ciphertext;
      assert.ok(ciphertext instanceof Uint8Array || ciphertext instanceof ArrayBuffer);
      const ciphertextBytes = ciphertext instanceof Uint8Array ? ciphertext : new Uint8Array(ciphertext);
      assert.equal(Buffer.from(ciphertextBytes).includes(Buffer.from(secret)), false);

      assert.ok(calls.some((c) => c.cmd === "power_query_cache_key_get_or_create"));

      // Close DB handles before deleting to keep fake-indexeddb happy.
      const db = await store.store.open();
      db.close();
    } finally {
      await deleteDatabase(DB_NAME).catch(() => {});
      globalThis.__TAURI__ = originalTauri;
      if (originalIndexedDB === undefined) {
        delete globalThis.indexedDB;
      } else {
        globalThis.indexedDB = originalIndexedDB;
      }
      if (originalIDBKeyRange === undefined) {
        delete globalThis.IDBKeyRange;
      } else {
        globalThis.IDBKeyRange = originalIDBKeyRange;
      }
    }
  },
);
