import { describe, expect, it, afterEach } from "vitest";

import { createDefaultAIAuditStore } from "../src/factory.node.js";
import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";

const originalGlobals = {
  indexedDB: Object.getOwnPropertyDescriptor(globalThis as any, "indexedDB"),
  IDBKeyRange: Object.getOwnPropertyDescriptor(globalThis as any, "IDBKeyRange"),
  localStorage: Object.getOwnPropertyDescriptor(globalThis as any, "localStorage"),
};

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

function unwrap(store: unknown): unknown {
  if (!(store instanceof BoundedAIAuditStore)) return store;
  // `store` is a private field but is present at runtime.
  return (store as any).store;
}

async function closeIfSupported(store: unknown): Promise<void> {
  const inner = unwrap(store) as any;
  if (inner && typeof inner.close === "function") {
    await inner.close();
  }
}

afterEach(async () => {
  if (originalGlobals.indexedDB) {
    Object.defineProperty(globalThis as any, "indexedDB", originalGlobals.indexedDB);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
  }

  if (originalGlobals.IDBKeyRange) {
    Object.defineProperty(globalThis as any, "IDBKeyRange", originalGlobals.IDBKeyRange);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;
  }

  if (originalGlobals.localStorage) {
    Object.defineProperty(globalThis as any, "localStorage", originalGlobals.localStorage);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).localStorage;
  }
});

describe("createDefaultAIAuditStore (node entrypoint)", () => {
  it("defaults to BoundedAIAuditStore wrapping MemoryAIAuditStore", async () => {
    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("propagates bounded options to BoundedAIAuditStore", async () => {
    const store = await createDefaultAIAuditStore({ bounded: { max_entry_chars: 123 } });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect((store as BoundedAIAuditStore).maxEntryChars).toBe(123);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("still defaults to memory even when indexedDB globals exist", async () => {
    Object.defineProperty(globalThis as any, "indexedDB", { value: indexedDB, configurable: true });
    Object.defineProperty(globalThis as any, "IDBKeyRange", { value: IDBKeyRange, configurable: true });

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('prefer: "localstorage" chooses LocalStorageAIAuditStore when localStorage is available', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis as any, "localStorage", { value: storage, configurable: true });

    const store = await createDefaultAIAuditStore({ prefer: "localstorage" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it('prefer: "localstorage" falls back to memory when localStorage access throws', async () => {
    Object.defineProperty(globalThis as any, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });

    const store = await createDefaultAIAuditStore({ prefer: "localstorage" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('prefer: "sqlite" returns a bounded wrapper by default', async () => {
    const store = await createDefaultAIAuditStore({ prefer: "sqlite" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(SqliteAIAuditStore);
    await closeIfSupported(store);
  });

  it('prefer: "sqlite" respects bounded:false', async () => {
    const store = await createDefaultAIAuditStore({ prefer: "sqlite", bounded: false });
    expect(store).toBeInstanceOf(SqliteAIAuditStore);
    await closeIfSupported(store);
  });
});
