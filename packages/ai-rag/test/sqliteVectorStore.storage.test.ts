// @vitest-environment jsdom

import { afterAll, beforeAll, beforeEach, expect, test } from "vitest";

import { LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";
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

maybeTest("SqliteVectorStore.compact() VACUUMs and persists a smaller DB", async () => {
  const storage = new LocalStorageBinaryStorage({
    namespace: "formula.test.rag.sqlite",
    workbookId: "sqlite-store-compact",
  });

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });

  const payload = "x".repeat(1024);
  const total = 400;
  const records = Array.from({ length: total }, (_, i) => ({
    id: `rec-${i}`,
    vector: i % 2 === 0 ? [1, 0, 0] : [0, 1, 0],
    metadata: { workbookId: "wb", i, payload },
  }));

  await store1.upsert(records);

  // Delete most records so the DB accumulates free pages.
  const remaining = 10;
  const deleteIds = Array.from({ length: total - remaining }, (_, i) => `rec-${i}`);
  await store1.delete(deleteIds);

  const before = (await storage.load())?.byteLength ?? 0;

  // Compact and ensure the store is still usable (dot() function still registered).
  await store1.compact();

  const after = (await storage.load())?.byteLength ?? 0;
  expect(after).toBeGreaterThan(0);
  expect(after).toBeLessThan(before);

  const remainingId = `rec-${total - 1}`;
  const rec = await store1.get(remainingId);
  expect(rec?.metadata?.i).toBe(total - 1);

  const hits = await store1.query([1, 0, 0], 5, { workbookId: "wb" });
  expect(hits.length).toBeGreaterThan(0);

  await store1.close();

  // Reload from persisted storage and ensure queries still work.
  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const rec2 = await store2.get(remainingId);
  expect(rec2?.metadata?.i).toBe(total - 1);

  const hits2 = await store2.query([1, 0, 0], 5, { workbookId: "wb" });
  expect(hits2.length).toBeGreaterThan(0);
  await store2.close();
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
