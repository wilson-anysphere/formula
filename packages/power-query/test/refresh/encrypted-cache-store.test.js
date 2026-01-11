import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { EncryptedFileSystemCacheStore } from "../../src/cache/encryptedFilesystem.js";
import { fnv1a64 } from "../../src/cache/key.js";

import { isEncryptedFileBytes } from "../../../security/crypto/encryptedFile.js";
import { InMemoryKeychainProvider } from "../../../security/crypto/keychain/inMemoryKeychain.js";

test("EncryptedFileSystemCacheStore: encrypts at rest and supports disabling encryption", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "pq-encrypted-cache-"));
  try {
    const keychainProvider = new InMemoryKeychainProvider();

    /** @type {import("../../src/cache/cache.js").CacheEntry} */
    const entry = { value: { hello: "world" }, createdAtMs: 1, expiresAtMs: null };
    const key = "cache-key-1";

    const store = new EncryptedFileSystemCacheStore({
      directory: dir,
      encryption: { enabled: true, keychainProvider }
    });
    await store.set(key, entry);

    const filePath = path.join(dir, `${fnv1a64(key)}.json`);
    const encryptedBytes = await fs.readFile(filePath);
    assert.equal(isEncryptedFileBytes(encryptedBytes), true);
    assert.throws(() => JSON.parse(encryptedBytes.toString("utf8")));

    const reloaded = new EncryptedFileSystemCacheStore({
      directory: dir,
      encryption: { enabled: true, keychainProvider }
    });
    assert.deepEqual(await reloaded.get(key), entry);

    await reloaded.disableEncryption();

    const plaintextBytes = await fs.readFile(filePath);
    assert.equal(isEncryptedFileBytes(plaintextBytes), false);
    assert.deepEqual(JSON.parse(plaintextBytes.toString("utf8")), { key, entry });
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

