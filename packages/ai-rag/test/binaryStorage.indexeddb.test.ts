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

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
  });
}

function transactionToPromise(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
  });
}

async function listKeys(dbName: string, storeName: string): Promise<IDBValidKey[]> {
  const req = fakeIndexedDB.open(dbName, 1);
  req.onupgradeneeded = () => {
    const db = req.result;
    if (!db.objectStoreNames.contains(storeName)) db.createObjectStore(storeName);
  };
  const db = await requestToPromise(req);
  try {
    const tx = db.transaction(storeName, "readonly");
    const store = tx.objectStore(storeName);
    const keys = await requestToPromise(store.getAllKeys());
    await transactionToPromise(tx);
    return keys;
  } finally {
    db.close();
  }
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

test("IndexedDBBinaryStorage namespaces per namespace", async () => {
  const dbName = uniqueDbName("ai-rag-idb-namespace2");

  const storageA = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.ns-a",
    workbookId: "wb-shared",
  });
  const storageB = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag.ns-b",
    workbookId: "wb-shared",
  });

  await storageA.save(new Uint8Array([1]));
  await storageB.save(new Uint8Array([2]));

  expect(Array.from((await storageA.load()) ?? [])).toEqual([1]);
  expect(Array.from((await storageB.load()) ?? [])).toEqual([2]);
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

test("IndexedDBBinaryStorage remove deletes persisted bytes", async () => {
  const dbName = uniqueDbName("ai-rag-idb-remove");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-remove",
  });

  const bytes = new Uint8Array([9, 8, 7]);
  await storage.save(bytes);
  expect(await listKeys(dbName, "binary")).toContain(storage.key);

  await storage.remove();
  expect(await storage.load()).toBeNull();
  expect(await listKeys(dbName, "binary")).not.toContain(storage.key);
});

test("IndexedDBBinaryStorage clears corrupted records on load", async () => {
  const dbName = uniqueDbName("ai-rag-idb-corrupt");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-corrupt",
  });

  // Write an invalid value directly into IndexedDB to simulate storage corruption.
  const req = fakeIndexedDB.open(dbName, 1);
  req.onupgradeneeded = () => {
    const db = req.result;
    if (!db.objectStoreNames.contains("binary")) db.createObjectStore("binary");
  };
  const db = await requestToPromise(req);
  try {
    const tx = db.transaction("binary", "readwrite");
    tx.objectStore("binary").put("not-bytes", storage.key);
    await transactionToPromise(tx);
  } finally {
    db.close();
  }

  expect(await listKeys(dbName, "binary")).toContain(storage.key);
  await expect(storage.load()).resolves.toBeNull();
  expect(await listKeys(dbName, "binary")).not.toContain(storage.key);
});

test("IndexedDBBinaryStorage remove is safe when key is missing", async () => {
  const dbName = uniqueDbName("ai-rag-idb-remove-missing");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-missing",
  });

  await expect(storage.remove()).resolves.toBeUndefined();
  expect(await storage.load()).toBeNull();
});

test("IndexedDBBinaryStorage retries opening after a transient open failure", async () => {
  const dbName = uniqueDbName("ai-rag-idb-retry");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-retry",
  });

  // Fail the first open attempt.
  Object.defineProperty(globalThis, "indexedDB", {
    value: {
      open() {
        throw new Error("boom");
      },
    },
    configurable: true,
  });

  await expect(storage.save(new Uint8Array([1, 2, 3]))).resolves.toBeUndefined();
  await expect(storage.load()).resolves.toBeNull();

  // Restore real IndexedDB implementation and ensure the storage recovers.
  Object.defineProperty(globalThis, "indexedDB", { value: fakeIndexedDB, configurable: true });

  await storage.save(new Uint8Array([1, 2, 3]));
  expect(Array.from((await storage.load()) ?? [])).toEqual([1, 2, 3]);
});

test("IndexedDBBinaryStorage can load Blob values (legacy storage shape)", async () => {
  const dbName = uniqueDbName("ai-rag-idb-blob");
  const storage = new IndexedDBBinaryStorage({
    dbName,
    namespace: "formula.test.rag",
    workbookId: "wb-blob",
  });

  // Seed the DB so the object store exists, then overwrite the stored value with a Blob.
  await storage.save(new Uint8Array([0]));

  const blob = new Blob([new Uint8Array([1, 2, 3, 4])]);
  const req = fakeIndexedDB.open(dbName, 1);
  const db = await requestToPromise(req);
  try {
    const tx = db.transaction("binary", "readwrite");
    tx.objectStore("binary").put(blob, storage.key);
    await transactionToPromise(tx);
  } finally {
    db.close();
  }

  expect(Array.from((await storage.load()) ?? [])).toEqual([1, 2, 3, 4]);
});
