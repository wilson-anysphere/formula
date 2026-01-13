import { afterAll, beforeAll, expect, test } from "vitest";

import { indexedDB as fakeIndexedDB } from "fake-indexeddb";

import { IndexedDBBinaryStorage } from "../src/store/binaryStorage.js";

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

test("IndexedDBBinaryStorage round-trips bytes", async () => {
  const dbName = uniqueDbName("ai-rag-idb-roundtrip");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-123",
  });

  expect(storage.key).toContain("wb-123");

  const bytes = new Uint8Array([1, 2, 3, 4, 255]);
  await storage.save(bytes);

  const loaded = await storage.load();
  expect(loaded).toBeInstanceOf(Uint8Array);
  expect(Array.from(loaded ?? [])).toEqual(Array.from(bytes));
});

test("IndexedDBBinaryStorage namespaces per workbookId", async () => {
  const dbName = uniqueDbName("ai-rag-idb-namespace");

  const storageA = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-a",
  });
  const storageB = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-b",
  });

  await storageA.save(new Uint8Array([1, 1, 1]));
  await storageB.save(new Uint8Array([2, 2, 2]));

  expect(Array.from((await storageA.load()) ?? [])).toEqual([1, 1, 1]);
  expect(Array.from((await storageB.load()) ?? [])).toEqual([2, 2, 2]);
});

test("IndexedDBBinaryStorage stores the exact Uint8Array view (respects byteOffset)", async () => {
  const dbName = uniqueDbName("ai-rag-idb-byteoffset");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-view",
  });

  const base = new Uint8Array([9, 9, 1, 2, 3, 9, 9]);
  const view = base.subarray(2, 5); // [1, 2, 3]
  await storage.save(view);

  expect(Array.from((await storage.load()) ?? [])).toEqual([1, 2, 3]);
});

test("IndexedDBBinaryStorage gracefully falls back when indexedDB is unavailable", async () => {
  const dbName = uniqueDbName("ai-rag-idb-fallback");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-fallback",
  });

  // Simulate an environment where IndexedDB is unavailable (e.g. Node, restricted webviews).
  Object.defineProperty(globalThis, "indexedDB", { value: undefined, configurable: true });
  try {
    await expect(storage.save(new Uint8Array([1, 2, 3]))).resolves.toBeUndefined();
    await expect(storage.load()).resolves.toBeNull();
  } finally {
    Object.defineProperty(globalThis, "indexedDB", { value: fakeIndexedDB, configurable: true });
  }

  // The earlier save should have been a no-op.
  expect(await storage.load()).toBeNull();
});
