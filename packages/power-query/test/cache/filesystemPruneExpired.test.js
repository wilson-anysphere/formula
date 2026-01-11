import assert from "node:assert/strict";
import { mkdtemp, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { FileSystemCacheStore } from "../../src/cache/filesystem.js";

async function exists(p) {
  try {
    await stat(p);
    return true;
  } catch {
    return false;
  }
}

test("FileSystemCacheStore.pruneExpired removes expired entries (and .bin blobs)", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-prune-fs-"));

  try {
    let now = 0;
    const store = new FileSystemCacheStore({ directory: cacheDir });
    const cache = new CacheManager({ store, now: () => now });

    await cache.set("expired", { bytes: new Uint8Array([1, 2, 3]) }, { ttlMs: 10 });
    await cache.set("alive", { ok: true }, { ttlMs: 100 });

    const expiredPaths = await store.pathsForKey("expired");
    const alivePaths = await store.pathsForKey("alive");

    assert.equal(await exists(expiredPaths.jsonPath), true);
    assert.equal(await exists(expiredPaths.binPath), true, "binary values should create a .bin blob");
    assert.equal(await exists(alivePaths.jsonPath), true);

    now = 20;
    await cache.pruneExpired();

    assert.equal(await exists(expiredPaths.jsonPath), false);
    assert.equal(await exists(expiredPaths.binPath), false);
    assert.equal(await exists(alivePaths.jsonPath), true);
    assert.deepEqual(await cache.get("alive"), { ok: true });
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

