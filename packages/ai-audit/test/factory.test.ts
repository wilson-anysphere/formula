import { afterEach, describe, expect, it } from "vitest";
import { randomUUID } from "node:crypto";

import { createDefaultAIAuditStore } from "../src/factory.js";
import { createDefaultAIAuditStore as createDefaultAIAuditStoreNode } from "../src/factory.node.js";
import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { IndexedDbAIAuditStore } from "../src/indexeddb-store.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import { InMemoryBinaryStorage } from "../src/storage.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";
import type { AIAuditEntry } from "../src/types.js";

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

  it("uses top-level max_entries/max_age_ms over retention", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({
      bounded: false,
      max_entries: 7,
      max_age_ms: 1234,
      retention: { max_entries: 42, max_age_ms: 9999 }
    });

    expect(store).toBeInstanceOf(IndexedDbAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).maxEntries).toBe(7);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).maxAgeMs).toBe(1234);
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

  it('prefer: "localstorage" forwards max_age_ms to LocalStorageAIAuditStore', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });

    const store = await createDefaultAIAuditStore({ prefer: "localstorage", max_age_ms: 1234, bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
    expect((store as LocalStorageAIAuditStore).maxAgeMs).toBe(1234);
  });

  it('prefer: "localstorage" wraps LocalStorageAIAuditStore in BoundedAIAuditStore by default', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });

    const store = await createDefaultAIAuditStore({ prefer: "localstorage" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it('prefer: "localstorage" wraps MemoryAIAuditStore fallback in BoundedAIAuditStore by default when localStorage is unavailable', async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    // Ensure IndexedDB is also present so this test catches accidental IndexedDB fallback.
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ prefer: "localstorage" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('prefer: "memory" chooses MemoryAIAuditStore even when persistence APIs exist', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ prefer: "memory", max_entries: 7, max_age_ms: 1234, bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).maxEntries).toBe(7);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).maxAgeMs).toBe(1234);
  });

  it('prefer: "memory" wraps MemoryAIAuditStore in BoundedAIAuditStore by default', async () => {
    const store = await createDefaultAIAuditStore({ prefer: "memory" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('prefer: "localstorage" falls back to MemoryAIAuditStore when localStorage is unavailable (even if indexedDB exists)', async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ prefer: "localstorage", bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("falls back to LocalStorageAIAuditStore when IndexedDB is unavailable", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    // Ensure IndexedDB globals are absent.
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;

    const store = await createDefaultAIAuditStore({ bounded: false, retention: { max_entries: 42 } });

    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
    expect((store as LocalStorageAIAuditStore).maxEntries).toBe(42);
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

  it('prefer: "indexeddb" wraps IndexedDbAIAuditStore in BoundedAIAuditStore by default when available', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(IndexedDbAIAuditStore);
  });

  it('prefer: "indexeddb" wraps LocalStorageAIAuditStore fallback in BoundedAIAuditStore by default when open fails', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it('prefer: "indexeddb" wraps MemoryAIAuditStore fallback in BoundedAIAuditStore by default when localStorage is unavailable', async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb" });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('prefer: "indexeddb" falls back to LocalStorageAIAuditStore when IndexedDB open errors (onerror)', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        const request: any = { error: new Error("indexedDB.open failed") };
        // Trigger the `onerror` callback after the store has attached handlers.
        Promise.resolve()
          .then(() => request.onerror?.())
          .catch(() => {});
        return request;
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb", bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it('prefer: "indexeddb" falls back to LocalStorageAIAuditStore when IndexedDB open is blocked', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        const request: any = {};
        // Trigger the `onblocked` callback after the store has attached handlers.
        Promise.resolve()
          .then(() => request.onblocked?.())
          .catch(() => {});
        return request;
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb", bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it('prefer: "indexeddb" falls back to MemoryAIAuditStore when IndexedDB fails and localStorage is unavailable', async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore({ prefer: "indexeddb", bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("falls back to LocalStorageAIAuditStore when IndexedDB open fails by default", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore({ bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it("falls back to MemoryAIAuditStore when localStorage access throws", async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    // Ensure IndexedDB is unavailable so the factory attempts localStorage next.
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;

    const store = await createDefaultAIAuditStore({ bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("falls back to MemoryAIAuditStore when localStorage setItem throws", async () => {
    const storage: any = new MemoryLocalStorage();
    storage.setItem = () => {
      throw new Error("quota exceeded");
    };
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    // Ensure IndexedDB is unavailable so the factory attempts localStorage next.
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).indexedDB;
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).IDBKeyRange;

    const store = await createDefaultAIAuditStore({ bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
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

  it("wraps the default store in BoundedAIAuditStore by default", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(IndexedDbAIAuditStore);
  });

  it("wraps LocalStorageAIAuditStore fallback in BoundedAIAuditStore by default when IndexedDB open fails", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(LocalStorageAIAuditStore);
  });

  it("wraps MemoryAIAuditStore fallback in BoundedAIAuditStore by default when localStorage is unavailable and IndexedDB fails", async () => {
    const win: any = {};
    Object.defineProperty(win, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    Object.defineProperty(globalThis, "window", { value: win, configurable: true });
    (globalThis as any).indexedDB = {
      open() {
        throw new Error("indexedDB.open failed");
      }
    };

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("propagates bounded options to the default wrapper", async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "window", { value: { localStorage: storage }, configurable: true });
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStore({ bounded: { max_entry_chars: 123 } });
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect((store as BoundedAIAuditStore).maxEntryChars).toBe(123);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    expect((store as any).store).toBeInstanceOf(IndexedDbAIAuditStore);
  });

  it("defaults to MemoryAIAuditStore in Node runtimes (no window)", async () => {
    // Ensure we don't accidentally treat the test environment as browser-like.
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete (globalThis as any).window;

    const store = await createDefaultAIAuditStore({ bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('node entrypoint: prefer "indexeddb" falls back to LocalStorageAIAuditStore when localStorage is available', async () => {
    const storage = new MemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { value: storage, configurable: true });
    // Even if IndexedDB globals are present in Node (e.g. via fake-indexeddb),
    // the Node entrypoint should *not* attempt to use IndexedDB.
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStoreNode({ prefer: "indexeddb", max_entries: 7, bounded: false });
    expect(store).toBeInstanceOf(LocalStorageAIAuditStore);
    expect((store as LocalStorageAIAuditStore).maxEntries).toBe(7);
  });

  it('node entrypoint: prefer "indexeddb" falls back to MemoryAIAuditStore when localStorage access throws', async () => {
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      }
    });
    // Ensure IndexedDB globals are present so this test catches accidental IndexedDB usage.
    (globalThis as any).indexedDB = indexedDB;
    (globalThis as any).IDBKeyRange = IDBKeyRange;

    const store = await createDefaultAIAuditStoreNode({ prefer: "indexeddb", bounded: false });
    expect(store).toBeInstanceOf(MemoryAIAuditStore);
  });

  it('throws when prefer: "sqlite" is used in the browser/default entrypoint', async () => {
    await expect(createDefaultAIAuditStore({ prefer: "sqlite" })).rejects.toThrow(
      'createDefaultAIAuditStore(prefer: "sqlite") is not available in the default/browser entrypoint. Import SqliteAIAuditStore from "@formula/ai-audit/sqlite" instead.'
    );
  });

  it('node entrypoint: prefer "sqlite" returns SqliteAIAuditStore and applies retention', async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await createDefaultAIAuditStoreNode({
      prefer: "sqlite",
      sqlite_storage: storage,
      max_entries: 1,
      bounded: false
    });

    expect(store).toBeInstanceOf(SqliteAIAuditStore);

    const base = Date.now();
    const e1: AIAuditEntry = {
      id: randomUUID(),
      timestamp_ms: base,
      session_id: "session-sqlite",
      mode: "chat",
      input: { prompt: "older" },
      model: "unit-test-model",
      tool_calls: []
    };
    const e2: AIAuditEntry = {
      ...e1,
      id: randomUUID(),
      timestamp_ms: base + 1,
      input: { prompt: "newer" }
    };

    await store.logEntry(e1);
    await store.logEntry(e2);

    const entries = await store.listEntries({ session_id: e1.session_id });
    expect(entries.map((e) => e.id)).toEqual([e2.id]);
  });
});
