import assert from "node:assert/strict";
import { mkdtemp, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { InMemoryKeychainProvider } from "../../../security/crypto/keychain/inMemoryKeychain.js";

import { EncryptedFileSystemCacheStore } from "../../src/cache/encryptedFilesystem.js";

test("EncryptedFileSystemCacheStore.pruneExpired removes expired entries and blobs", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-prune-efs-"));
  try {
    const keychainProvider = new InMemoryKeychainProvider();
    const store = new EncryptedFileSystemCacheStore({ directory: dir, encryption: { enabled: true, keychainProvider } });

    await store.set("expired", {
      value: { version: 2, table: { kind: "arrow", format: "ipc", columns: [], bytes: new Uint8Array([1, 2, 3]) }, meta: null },
      createdAtMs: 0,
      expiresAtMs: 5,
    });
    await store.set("alive", { value: { ok: true }, createdAtMs: 0, expiresAtMs: 50 });

    const { jsonPath: expiredJson, binPath: expiredBin } = await store.pathsForKey("expired");
    await stat(expiredJson);
    await stat(expiredBin);

    await store.pruneExpired(10);

    assert.equal(await store.get("expired"), null);
    assert.ok(await store.get("alive"));
    await assert.rejects(stat(expiredJson));
    await assert.rejects(stat(expiredBin));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

