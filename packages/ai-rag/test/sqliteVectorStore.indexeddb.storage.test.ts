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

maybeTest("SqliteVectorStore throws typed mismatch error on dimension mismatch when reset is disabled (IndexedDBBinaryStorage)", async () => {
  const dbName = uniqueDbName("ai-rag-sqlite-idb-dim-mismatch-throw");
  const storage1 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch-throw",
  });

  const store1 = await SqliteVectorStore.create({ storage: storage1, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  await store1.close();

  const storage2 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch-throw",
  });

  let removed = false;
  const originalRemove = storage2.remove.bind(storage2);
  // eslint-disable-next-line no-param-reassign
  storage2.remove = async () => {
    removed = true;
    await originalRemove();
  };

  await expect(
    SqliteVectorStore.create({
      storage: storage2,
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

  // Ensure the store wasn't wiped.
  const storage3 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-dim-mismatch-throw",
  });
  const store3 = await SqliteVectorStore.create({
    storage: storage3,
    dimension: 3,
    autoSave: false,
    resetOnDimensionMismatch: false,
    resetOnCorrupt: false,
  });
  const rec = await store3.get("a");
  expect(rec).not.toBeNull();
  await store3.close();
});

maybeTest("SqliteVectorStore throws typed invalid metadata error on invalid dimension meta when reset is disabled (IndexedDBBinaryStorage)", async () => {
  const dbName = uniqueDbName("ai-rag-sqlite-idb-invalid-dim-meta");
  const storage1 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-invalid-dim-meta",
  });

  const store1 = await SqliteVectorStore.create({ storage: storage1, dimension: 3, autoSave: true });
  // Capture the underlying sql.js Database constructor so we can inspect persisted
  // bytes without re-initializing sql.js in this test.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const DatabaseCtor = (store1 as any)._db.constructor as any;

  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
  // Corrupt the persisted dimension metadata and persist it.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (store1 as any)._db.run("UPDATE vector_store_meta SET value = 'not-a-number' WHERE key = 'dimension';");
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (store1 as any)._dirty = true;
  await store1.close();

  const storage2 = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-invalid-dim-meta",
  });

  let removed = false;
  const originalRemove = storage2.remove.bind(storage2);
  // eslint-disable-next-line no-param-reassign
  storage2.remove = async () => {
    removed = true;
    await originalRemove();
  };

  await expect(
    SqliteVectorStore.create({
      storage: storage2,
      dimension: 4,
      autoSave: false,
      resetOnDimensionMismatch: false,
      resetOnCorrupt: false,
    })
  ).rejects.toMatchObject({
    name: "SqliteVectorStoreInvalidMetadataError",
    rawDimension: "not-a-number",
  });
  expect(removed).toBe(false);

  // Ensure the DB bytes are still present and still contain the corrupted meta + record.
  const bytes = await storage2.load();
  expect(bytes).not.toBeNull();

  const db = new DatabaseCtor(bytes);
  const metaStmt = db.prepare("SELECT value FROM vector_store_meta WHERE key = 'dimension' LIMIT 1;");
  expect(metaStmt.step()).toBe(true);
  expect(metaStmt.get()[0]).toBe("not-a-number");
  metaStmt.free();

  const countStmt = db.prepare("SELECT COUNT(*) FROM vectors;");
  expect(countStmt.step()).toBe(true);
  expect(Number(countStmt.get()[0])).toBe(1);
  countStmt.free();

  db.close();
});
