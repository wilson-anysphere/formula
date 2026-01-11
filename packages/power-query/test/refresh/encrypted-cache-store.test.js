import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { arrowTableFromColumns } from "../../../data-io/src/index.js";

import { CacheManager } from "../../src/cache/cache.js";
import { EncryptedFileSystemCacheStore } from "../../src/cache/encryptedFilesystem.js";
import { deserializeAnyTable, serializeAnyTable } from "../../src/cache/serialize.js";
import { ArrowTableAdapter } from "../../src/arrowTable.js";
import { QueryEngine } from "../../src/engine.js";
import { fnv1a64 } from "../../src/cache/key.js";

import { isEncryptedFileBytes } from "../../../security/crypto/encryptedFile.js";
import { InMemoryKeychainProvider } from "../../../security/crypto/keychain/inMemoryKeychain.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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

test("EncryptedFileSystemCacheStore: stores Arrow IPC payloads in an encrypted .bin blob", async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "pq-encrypted-cache-arrow-"));
  try {
    const keychainProvider = new InMemoryKeychainProvider();

    const adapter = new ArrowTableAdapter(
      arrowTableFromColumns({
        id: new Int32Array([1, 2]),
        name: ["Alice", "Bob"],
      }),
    );

    const cacheValue = { version: 2, table: serializeAnyTable(adapter), meta: null };

    /** @type {import("../../src/cache/cache.js").CacheEntry} */
    const entry = { value: cacheValue, createdAtMs: 1, expiresAtMs: null };
    const key = "cache-key-arrow";

    const store = new EncryptedFileSystemCacheStore({
      directory: dir,
      encryption: { enabled: true, keychainProvider },
    });
    await store.set(key, entry);

    const hash = fnv1a64(key);
    const jsonPath = path.join(dir, `${hash}.json`);
    const binPath = path.join(dir, `${hash}.bin`);

    const jsonBytes = await fs.readFile(jsonPath);
    const binBytes = await fs.readFile(binPath);
    assert.equal(isEncryptedFileBytes(jsonBytes), true);
    assert.equal(isEncryptedFileBytes(binBytes), true);

    const reloaded = new EncryptedFileSystemCacheStore({
      directory: dir,
      encryption: { enabled: true, keychainProvider },
    });
    const loaded = await reloaded.get(key);
    assert.ok(loaded);
    assert.ok(loaded.value?.table?.bytes instanceof Uint8Array);

    const roundTrip = deserializeAnyTable(loaded.value.table);
    assert.deepEqual(roundTrip.toGrid(), adapter.toGrid());

    await reloaded.disableEncryption();

    const jsonBytes2 = await fs.readFile(jsonPath);
    const binBytes2 = await fs.readFile(binPath);
    assert.equal(isEncryptedFileBytes(jsonBytes2), false);
    assert.equal(isEncryptedFileBytes(binBytes2), false);

    const parsed = JSON.parse(jsonBytes2.toString("utf8"));
    assert.equal(parsed.key, key);
    assert.ok(parsed.entry?.value?.table?.bytes?.__pq_cache_binary);

    const afterDisable = new EncryptedFileSystemCacheStore({
      directory: dir,
      encryption: { enabled: false, keychainProvider },
    });
    const loadedAfterDisable = await afterDisable.get(key);
    assert.ok(loadedAfterDisable);
    assert.deepEqual(deserializeAnyTable(loadedAfterDisable.value.table).toGrid(), adapter.toGrid());
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test("QueryEngine: caches Arrow-backed Parquet results using EncryptedFileSystemCacheStore", async () => {
  const cacheDir = await fs.mkdtemp(path.join(os.tmpdir(), "pq-encrypted-cache-engine-"));
  try {
    const keychainProvider = new InMemoryKeychainProvider();
    const parquetPath = path.join(__dirname, "..", "..", "..", "data-io", "test", "fixtures", "simple.parquet");

    let readCount = 0;
    const firstEngine = new QueryEngine({
      cache: new CacheManager({
        store: new EncryptedFileSystemCacheStore({ directory: cacheDir, encryption: { enabled: true, keychainProvider } }),
      }),
      fileAdapter: {
        readBinary: async (p) => {
          readCount += 1;
          return new Uint8Array(await fs.readFile(p));
        },
      },
    });

    const query = { id: "q_parquet_encrypted_fs_cache", name: "Parquet encrypted fs cache", source: { type: "parquet", path: parquetPath }, steps: [] };

    const first = await firstEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(first.meta.cache?.hit, false);
    assert.equal(readCount, 1);
    assert.ok(first.table instanceof ArrowTableAdapter);
    const grid = first.table.toGrid();

    let secondReadCount = 0;
    const secondEngine = new QueryEngine({
      cache: new CacheManager({
        store: new EncryptedFileSystemCacheStore({ directory: cacheDir, encryption: { enabled: true, keychainProvider } }),
      }),
      fileAdapter: {
        readBinary: async (p) => {
          secondReadCount += 1;
          return new Uint8Array(await fs.readFile(p));
        },
      },
    });

    const second = await secondEngine.executeQueryWithMeta(query, {}, {});
    assert.equal(second.meta.cache?.hit, true);
    assert.equal(secondReadCount, 0, "cache hit should not re-read Parquet bytes");
    assert.ok(second.table instanceof ArrowTableAdapter);
    assert.deepEqual(second.table.toGrid(), grid);
  } finally {
    await fs.rm(cacheDir, { recursive: true, force: true });
  }
});
