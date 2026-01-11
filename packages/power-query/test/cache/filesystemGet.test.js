import assert from "node:assert/strict";
import { mkdtemp, rm, stat, writeFile } from "node:fs/promises";
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

test("FileSystemCacheStore.get cleans up corrupted cache entries", async () => {
  const cacheDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-get-corrupt-"));

  try {
    const store = new FileSystemCacheStore({ directory: cacheDir });
    const cache = new CacheManager({ store });

    await cache.set("corrupt", { bytes: new Uint8Array([1, 2, 3]) });
    const { jsonPath, binPath } = await store.pathsForKey("corrupt");

    assert.equal(await exists(jsonPath), true);
    assert.equal(await exists(binPath), true);

    await writeFile(jsonPath, "{ not json", "utf8");
    assert.equal(await cache.get("corrupt"), null);

    assert.equal(await exists(jsonPath), false);
    assert.equal(await exists(binPath), false);
  } finally {
    await rm(cacheDir, { recursive: true, force: true });
  }
});

