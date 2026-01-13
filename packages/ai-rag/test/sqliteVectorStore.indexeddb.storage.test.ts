import { afterAll, beforeAll, expect, test } from "vitest";

import { indexedDB as fakeIndexedDB } from "fake-indexeddb";

import { IndexedDBBinaryStorage } from "../src/store/binaryStorage.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

const originalIndexedDB = Object.getOwnPropertyDescriptor(globalThis, "indexedDB");

beforeAll(() => {
  Object.defineProperty(globalThis, "indexedDB", { value: fakeIndexedDB, configurable: true });
});

afterAll(() => {
  if (originalIndexedDB) {
    Object.defineProperty(globalThis, "indexedDB", originalIndexedDB);
  }
});

function uniqueDbName(prefix: string): string {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

const maybeTest = sqlJsAvailable ? test : test.skip;

maybeTest("SqliteVectorStore persists and reloads via IndexedDBBinaryStorage", async () => {
  const dbName = uniqueDbName("ai-rag-sqlite-idb");
  const storage1 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store",
  });

  const store1 = await SqliteVectorStore.create({ storage: storage1, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
  ]);
  await store1.close();

  // Simulate restart (new storage instance, same key/db).
  const storage2 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store",
  });

  const store2 = await SqliteVectorStore.create({ storage: storage2, dimension: 3, autoSave: false });
  const rec = await store2.get("a");
  expect(rec).not.toBeNull();
  expect(rec?.metadata?.label).toBe("A");

  const hits = await store2.query([1, 0, 0], 1, { workbookId: "wb" });
  expect(hits[0]?.id).toBe("a");
  await store2.close();
});

maybeTest("SqliteVectorStore can reset persisted DB on dimension mismatch (IndexedDBBinaryStorage)", async () => {
  const dbName = uniqueDbName("ai-rag-sqlite-idb-dim-mismatch");
  const storage1 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch",
  });

  const store1 = await SqliteVectorStore.create({ storage: storage1, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  // Simulate restart with mismatched dimension. Reset should wipe the DB.
  const storage2 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch",
  });
  const store2 = await SqliteVectorStore.create({
    storage: storage2,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: true,
  });
  expect(await store2.list()).toEqual([]);
  await store2.close();

  // Ensure the reset wrote a compatible DB so future opens don't repeatedly reset.
  const storage3 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch",
  });
  const store3 = await SqliteVectorStore.create({
    storage: storage3,
    dimension: 4,
    autoSave: false,
    resetOnDimensionMismatch: false,
  });
  expect(await store3.list()).toEqual([]);
  await store3.close();
});
