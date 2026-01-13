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

