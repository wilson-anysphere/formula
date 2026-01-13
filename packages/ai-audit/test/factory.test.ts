import { afterEach, describe, expect, it } from "vitest";

import { createDefaultAIAuditStore } from "../src/factory.js";
import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { IndexedDbAIAuditStore } from "../src/indexeddb-store.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";

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

const originalWindowDescriptor = Object.getOwnPropertyDescriptor(globalThis, "window");
const originalLocalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalIndexedDbDescriptor = Object.getOwnPropertyDescriptor(globalThis, "indexedDB");
const originalIdbKeyRangeDescriptor = Object.getOwnPropertyDescriptor(globalThis, "IDBKeyRange");

function restoreGlobals() {
  if (originalWindowDescriptor) Object.defineProperty(globalThis, "window", originalWindowDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).window;
  }

  if (originalLocalStorageDescriptor) Object.defineProperty(globalThis, "localStorage", originalLocalStorageDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).localStorage;
  }

  if (originalIndexedDbDescriptor) Object.defineProperty(globalThis, "indexedDB", originalIndexedDbDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
  }

  if (originalIdbKeyRangeDescriptor) Object.defineProperty(globalThis, "IDBKeyRange", originalIdbKeyRangeDescriptor);
  else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;
  }
}

describe("createDefaultAIAuditStore", () => {
  afterEach(() => {
    restoreGlobals();
  });

  it("prefers IndexedDbAIAuditStore in a browser-like runtime when indexedDB is available", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ retention: { max_entries: 42 }, bounded: false });

    expect(store).toBeInstanceOf(IndexedDbAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).maxEntries).toBe(42);
  });

  it('prefer: "localstorage" chooses LocalStorageAIAuditStore even when indexedDB exists', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ prefer: "localstorage", retention: { max_entries: 7 }, bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
    expect((store as LocalStorageAIAuditStore).maxEntries).toBe(7);
  });

  it('prefer: "indexeddb" falls back to LocalStorageAIAuditStore when IndexedDB open fails', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb", bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it("wraps the selected store in BoundedAIAuditStore when bounded is enabled", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });

    const store = await createDefaultAIAuditStore({ prefer: "localstorage", bounded: { max_entry_chars: 123 } });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect((store as BoundedAIAuditStore).maxEntryChars).toBe(123);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(LocalStorageAIAuditStore);
  });
});

