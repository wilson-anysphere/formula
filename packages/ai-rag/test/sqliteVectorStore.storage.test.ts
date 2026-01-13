// @vitest-environment jsdom

import { afterAll, beforeAll, beforeEach, expect, test } from "vitest";

import { ChunkedLocalStorageBinaryStorage, InMemoryBinaryStorage, LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";
import { ensureTestLocalStorage } from "./testLocalStorage.js";

ensureTestLocalStorage();

function getTestLocalStorage(): Storage {
  const jsdomStorage = (globalThis as any)?.jsdom?.window?.localStorage as Storage | undefined;
  if (!jsdomStorage) {
    throw new Error("Expected vitest jsdom environment to provide globalThis.jsdom.window.localStorage");
  }
  return jsdomStorage;
}

class MemoryLocalStorage implements Storage {
  #data = new Map<string, string>();

  get length(): number {
    return this.#data.size;
  }

  clear(): void {
    this.#data.clear();
  }

  getItem(key: string): string | null {
    return this.#data.get(key) ?? null;
  }

  key(index: number): string | null {
    return Array.from(this.#data.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.#data.delete(key);
  }

  setItem(key: string, value: string): void {
    this.#data.set(key, value);
  }
}

const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

beforeAll(() => {
  Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });
});

afterAll(() => {
  if (originalLocalStorage) {
    Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
  }
});

beforeEach(() => {
  getTestLocalStorage().clear();
});

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

const maybeTest = sqlJsAvailable ? test : test.skip;

maybeTest("SqliteVectorStore persists and reloads via BinaryStorage", async () => {
  const storage = new LocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store",
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
  ]);
  await store1.close();

  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  expect(rec?.metadata?.label).toBe("A");

  const hits = await store2.query([1, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("a");
  await store2.close();
});

maybeTest("SqliteVectorStore persists and reloads via ChunkedLocalStorageBinaryStorage", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-chunked",
    chunkSizeChars: 64,
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
  ]);
  await store1.close();

  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  expect(rec?.metadata?.label).toBe("A");

  const hits = await store2.query([1, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("a");
  await store2.close();
});

maybeTest("SqliteVectorStore can reset persisted DB on dimension mismatch", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: true,
    resetOnDimensionMismatch: true,
  });

  expect(await store2.list()).toEqual([]);

  await store2.upsert([{ id: "c", vector: [1, 0, 0, 0], metadata: { workbookId: "wb" } }]);
  const hits = await store2.query([1, 0, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("c");

  await store2.close();
});

maybeTest("SqliteVectorStore resets dimension mismatch by default (resetOnDimensionMismatch=true)", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  // Do not pass resetOnDimensionMismatch; default should be true.
  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: true,
  });

  expect(await store2.list()).toEqual([]);
  await store2.close();
});

maybeTest("SqliteVectorStore calls storage.remove() when resetting dimension mismatch", async () => {
  class SpyStorage {
    removed = false;
    #data: Uint8Array | null = null;

    async load(): Promise<Uint8Array | null> {
      return this.#data ? new Uint8Array(this.#data) : null;
    }

    async save(data: Uint8Array): Promise<void> {
      this.#data = new Uint8Array(data);
    }

    async remove(): Promise<void> {
      this.removed = true;
      this.#data = null;
    }
  }

  const storage = new SpyStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  expect(storage.removed).toBe(false);

  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: true,
  });

  expect(storage.removed).toBe(true);
  expect(await store2.list()).toEqual([]);
  await store2.close();
});

maybeTest("SqliteVectorStore treats invalid dimension metadata as corruption (not mismatch)", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);

  // Corrupt the persisted dimension metadata.
  (store1 as any)._db.run("UPDATE vector_store_meta SET value = 'not-a-number' WHERE key = 'dimension';");
  (store1 as any)._dirty = true;
  await store1.close();

  // Even with resetOnDimensionMismatch disabled, this should be treated as corruption
  // and recover as an empty store when resetOnCorrupt is enabled.
  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: false,
    resetOnCorrupt: true,
  });
  expect(await store2.list()).toEqual([]);
  await store2.close();
});

maybeTest("SqliteVectorStore throws on invalid dimension metadata when resetOnCorrupt=false", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);

  (store1 as any)._db.run("UPDATE vector_store_meta SET value = 'not-a-number' WHERE key = 'dimension';");
  (store1 as any)._dirty = true;
  await store1.close();

  let err: any = null;
  try {
    await SqliteVectorStore.create({
      storage,
      dimension: 4,
      autoSave: false,
      resetOnDimensionMismatch: false,
      resetOnCorrupt: false,
    });
  } catch (e) {
    err = e;
  }

  expect(err).toBeTruthy();
  expect(err).toMatchObject({
    name: "SqliteVectorStoreInvalidMetadataError",
    rawDimension: "not-a-number",
  });
});

maybeTest("SqliteVectorStore can reset dimension mismatch even when storage.remove() is missing", async () => {
  class NoRemoveBinaryStorage {
    #data: Uint8Array | null = null;

    async load(): Promise<Uint8Array | null> {
      return this.#data ? new Uint8Array(this.#data) : null;
    }

    async save(data: Uint8Array): Promise<void> {
      this.#data = new Uint8Array(data);
    }
  }

  const storage = new NoRemoveBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  // Reset a mismatched-dimension DB without relying on `storage.remove()`.
  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: true,
  });
  expect(await store2.list()).toEqual([]);

  // `create()` should have already persisted the fresh DB. If it didn't, this
  // would re-open the old mismatched DB and throw.
  const store3 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: false,
  });
  expect(await store3.list()).toEqual([]);

  await store2.close();
  await store3.close();
});

maybeTest("SqliteVectorStore can reset dimension mismatch when using ChunkedLocalStorageBinaryStorage", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-chunked-dim-mismatch",
    chunkSizeChars: 128,
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: true,
    resetOnDimensionMismatch: true,
  });

  expect(await store2.list()).toEqual([]);

  await store2.upsert([{ id: "c", vector: [1, 0, 0, 0], metadata: { workbookId: "wb" } }]);
  const hits = await store2.query([1, 0, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("c");

  await store2.close();

  // Ensure subsequent opens do not repeatedly reset.
  const store3 = await SqliteVectorStore.create({
    storage,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: false,
  });
  const list = await store3.list();
  expect(list.map((r) => r.id)).toEqual(["c"]);
  await store3.close();
});

function listKeysWithPrefix(storage: Storage, prefix: string): string[] {
  const keys: string[] = [];
  for (let i = 0; i < storage.length; i += 1) {
    const key = storage.key(i);
    if (key?.startsWith(prefix)) keys.push(key);
  }
  return keys.sort();
}

maybeTest("SqliteVectorStore does not reset chunked localStorage on dimension mismatch when reset is disabled", async () => {
  const storage = new ChunkedLocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-chunked-dim-mismatch-no-reset",
    chunkSizeChars: 64,
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  const localStorage = getTestLocalStorage();
  const beforeKeys = listKeysWithPrefix(localStorage, `${storage.key}:`);
  expect(beforeKeys.length).toBeGreaterThan(0);

  await expect(
    SqliteVectorStore.create({
      storage,
      dimension: 4,
      autoSave: false,
      resetOnDimensionMismatch: false,
      resetOnCorrupt: false,
    })
  ).rejects.toMatchObject({
    name: "SqliteVectorStoreDimensionMismatchError",
    dbDimension: 3,
    requestedDimension: 4,
  });

  const afterKeys = listKeysWithPrefix(localStorage, `${storage.key}:`);
  expect(afterKeys).toEqual(beforeKeys);

  // Ensure the original DB is still readable.
  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 3,
    autoSave: false,
    resetOnDimensionMismatch: false,
    resetOnCorrupt: false,
  });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  await store2.close();
});

maybeTest("SqliteVectorStore does not reset LocalStorageBinaryStorage on dimension mismatch when reset is disabled", async () => {
  const storage = new LocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-local-dim-mismatch-no-reset",
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  const localStorage = getTestLocalStorage();
  const before = localStorage.getItem(storage.key);
  expect(before).toBeTypeOf("string");

  let removed = false;
  const originalRemove = (storage as any).remove?.bind(storage) as (() => Promise<void>) | undefined;
  // eslint-disable-next-line no-param-reassign
  (storage as any).remove = async () => {
    removed = true;
    if (originalRemove) await originalRemove();
  };

  await expect(
    SqliteVectorStore.create({
      storage,
      dimension: 4,
      autoSave: false,
      resetOnDimensionMismatch: false,
      resetOnCorrupt: false,
    })
  ).rejects.toMatchObject({
    name: "SqliteVectorStoreDimensionMismatchError",
    dbDimension: 3,
    requestedDimension: 4,
  });
  expect(removed).toBe(false);

  const after = localStorage.getItem(storage.key);
  expect(after).toBe(before);

  // Ensure the original DB is still readable.
  const store2 = await SqliteVectorStore.create({
    storage,
    dimension: 3,
    autoSave: false,
    resetOnDimensionMismatch: false,
    resetOnCorrupt: false,
  });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  await store2.close();
});

maybeTest("SqliteVectorStore throws on dimension mismatch when resetOnDimensionMismatch=false", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  let err: any = null;
  try {
    await SqliteVectorStore.create({ storage, dimension: 4, autoSave: false, resetOnDimensionMismatch: false });
  } catch (e) {
    err = e;
  }
  expect(err).toBeTruthy();
  expect(err).toMatchObject({
    name: "SqliteVectorStoreDimensionMismatchError",
    dbDimension: 3,
    requestedDimension: 4,
  });
});

maybeTest("SqliteVectorStore.compact() VACUUMs and persists a smaller DB (even with autoSave:false)", async () => {
  const storage = new LocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-compact",
  });

  // Create a large DB snapshot.
  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });

  const payload = "x".repeat(1024);
  const total = 400;
  const records = Array.from({ length: total }, (_, i) => ({
    id: `rec-${i}`,
    vector: i % 2 === 0 ? [1, 0, 0] : [0, 1, 0],
    metadata: { workbookId: "wb", i, payload },
  }));

  await store1.upsert(records);
  // Persist the initial snapshot. (`autoSave:false` skips persisting on upsert.)
  await store1.close();

  // Reopen, delete most records so the DB accumulates free pages, and persist the
  // post-delete (but un-compacted) snapshot.
  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const remaining = 10;
  const deleteIds = Array.from({ length: total - remaining }, (_, i) => `rec-${i}`);
  await store2.delete(deleteIds);
  await store2.close();

  const before = (await storage.load())?.byteLength ?? 0;
  expect(before).toBeGreaterThan(0);

  // Compact and ensure the store is still usable (dot() function still registered).
  const store3 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  await store3.compact();

  const after = (await storage.load())?.byteLength ?? 0;
  expect(after).toBeGreaterThan(0);
  expect(after).toBeLessThan(before);

  const remainingId = `rec-${total - 1}`;
  const rec = await store3.get(remainingId);
  expect(rec?.metadata?.i).toBe(total - 1);

  const hits = await store3.query([1, 0, 0], 5, { workbookId: "wb" });
  expect(hits.length).toBeGreaterThan(0);

  await store3.close();

  // Reload from persisted storage and ensure queries still work.
  const store4 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const rec2 = await store4.get(remainingId);
  expect(rec2?.metadata?.i).toBe(total - 1);

  const hits2 = await store4.query([1, 0, 0], 5, { workbookId: "wb" });
  expect(hits2.length).toBeGreaterThan(0);
  await store4.close();
});

class CountingBinaryStorage {
  saveCalls = 0;
  #data: Uint8Array | null = null;

  async load(): Promise<Uint8Array | null> {
    return this.#data ? new Uint8Array(this.#data) : null;
  }

  async save(data: Uint8Array): Promise<void> {
    this.saveCalls += 1;
    this.#data = new Uint8Array(data);
  }
}

maybeTest("SqliteVectorStore does not persist on close when nothing changed", async () => {
  const storage = new CountingBinaryStorage();
  const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store.close();
  expect(storage.saveCalls).toBe(0);
});

maybeTest("SqliteVectorStore avoids double-persist on close after autoSave upsert", async () => {
  const storage = new CountingBinaryStorage();
  const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store.upsert([{ id: "a", vector: [1, 0, 0], metadata: {} }]);
  expect(storage.saveCalls).toBe(1);
  await store.close();
  expect(storage.saveCalls).toBe(1);
});
