import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { performance } from "node:perf_hooks";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { ArrowTableAdapter } from "../src/arrowTable.js";
import { CacheManager } from "../src/cache/cache.js";
import { EncryptedFileSystemCacheStore } from "../src/cache/encryptedFilesystem.js";
import { FileSystemCacheStore } from "../src/cache/filesystem.js";
import { MemoryCacheStore } from "../src/cache/memory.js";
import { deserializeAnyTable, serializeAnyTable } from "../src/cache/serialize.js";

import { InMemoryKeychainProvider } from "../../security/crypto/keychain/inMemoryKeychain.js";

function fmtMs(ms) {
  return `${ms.toFixed(1)}ms`;
}

function mem() {
  const { heapUsed } = process.memoryUsage();
  return `${(heapUsed / 1024 / 1024).toFixed(1)}MB`;
}

/**
 * @param {number} rowCount
 */
function makeArrowTable(rowCount) {
  const region = new Array(rowCount);
  const sales = new Float64Array(rowCount);
  const id = new Int32Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    region[i] = i & 1 ? "East" : "West";
    sales[i] = (i % 10_000) * 0.5;
    id[i] = i;
  }

  return new ArrowTableAdapter(
    arrowTableFromColumns({
      id,
      region,
      sales,
    }),
  );
}

async function benchStore(label, store, value, payloadBytes) {
  const cache = new CacheManager({ store });

  const setStart = performance.now();
  await cache.set("bench:key", value);
  const setEnd = performance.now();

  const getStart = performance.now();
  const cached = await cache.get("bench:key");
  const getEnd = performance.now();

  const deserStart = performance.now();
  const table = deserializeAnyTable(cached.table);
  const deserEnd = performance.now();

  // Touch output to avoid the work being optimized away.
  void table.head(1).toGrid();

  console.log(
    `${label}: set=${fmtMs(setEnd - setStart)} get=${fmtMs(getEnd - getStart)} deserialize=${fmtMs(deserEnd - deserStart)} bytes=${payloadBytes.toLocaleString()}`,
  );
}

const ROWS = 1_000_000;

console.log("Power Query Arrow cache benchmark");
console.log(`Node ${process.version}`);
console.log(`rows=${ROWS.toLocaleString()}`);
console.log("");

console.log(`create table... (heap=${mem()})`);
const createStart = performance.now();
const table = makeArrowTable(ROWS);
const createEnd = performance.now();
console.log(`create=${fmtMs(createEnd - createStart)} rowCount=${table.rowCount} heap=${mem()}`);
console.log("");

console.log("serialize (Arrow IPC)...");
const serStart = performance.now();
const tablePayload = serializeAnyTable(table);
const serEnd = performance.now();
console.log(`serialize=${fmtMs(serEnd - serStart)} kind=${tablePayload.kind} bytes=${tablePayload.bytes.byteLength.toLocaleString()} heap=${mem()}`);
console.log("");

console.log("deserialize (Arrow IPC)...");
const deserStart = performance.now();
const roundTrip = deserializeAnyTable(tablePayload);
const deserEnd = performance.now();
void roundTrip.head(1).toGrid();
console.log(`deserialize=${fmtMs(deserEnd - deserStart)} heap=${mem()}`);
console.log("");

const cacheValue = { version: 2, table: tablePayload, meta: null };

console.log("CacheStore benchmarks");
await benchStore("MemoryCacheStore", new MemoryCacheStore(), cacheValue, tablePayload.bytes.byteLength);

const tmpDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-arrow-"));
try {
  await benchStore(
    "FileSystemCacheStore",
    new FileSystemCacheStore({ directory: tmpDir }),
    cacheValue,
    tablePayload.bytes.byteLength,
  );
} finally {
  await rm(tmpDir, { recursive: true, force: true });
}

const tmpEncryptedDir = await mkdtemp(path.join(os.tmpdir(), "pq-cache-arrow-encrypted-"));
try {
  const keychainProvider = new InMemoryKeychainProvider();
  await benchStore(
    "EncryptedFileSystemCacheStore (plaintext)",
    new EncryptedFileSystemCacheStore({ directory: tmpEncryptedDir, encryption: { enabled: false, keychainProvider } }),
    cacheValue,
    tablePayload.bytes.byteLength,
  );
  await benchStore(
    "EncryptedFileSystemCacheStore (encrypted)",
    new EncryptedFileSystemCacheStore({ directory: tmpEncryptedDir, encryption: { enabled: true, keychainProvider } }),
    cacheValue,
    tablePayload.bytes.byteLength,
  );
} finally {
  await rm(tmpEncryptedDir, { recursive: true, force: true });
}
