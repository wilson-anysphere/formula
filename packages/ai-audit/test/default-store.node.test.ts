import { IDBKeyRange, indexedDB } from "fake-indexeddb";
import { describe, expect, it, afterEach } from "vitest";

import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { createDefaultAIAuditStore } from "../src/index.js";
import { IndexedDbAIAuditStore } from "../src/indexeddb-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";

const originalGlobals = {
  indexedDB: Object.getOwnPropertyDescriptor(globalThis as any, "indexedDB"),
  IDBKeyRange: Object.getOwnPropertyDescriptor(globalThis as any, "IDBKeyRange"),
  nodeVersion: Object.getOwnPropertyDescriptor(process.versions, "node"),
};

afterEach(() => {
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

  if (originalGlobals.nodeVersion) {
    Object.defineProperty(process.versions, "node", originalGlobals.nodeVersion);
  }
});

function unwrap(store: unknown): unknown {
  if (!(store instanceof BoundedAIAuditStore)) return store;
  // `store` is a private field but is present at runtime.
  return (store as any).store;
}

describe("createDefaultAIAuditStore (node)", () => {
  it("defaults to MemoryAIAuditStore when no browser storage is available", async () => {
    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("chooses IndexedDbAIAuditStore when indexedDB is present (fake-indexeddb)", async () => {
    if (originalGlobals.nodeVersion) {
      Object.defineProperty(process.versions, "node", { value: undefined, configurable: true });
    }
    Object.defineProperty(globalThis as any, "indexedDB", { value: indexedDB, configurable: true });
    Object.defineProperty(globalThis as any, "IDBKeyRange", { value: IDBKeyRange, configurable: true });

    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(IndexedDbAIAuditStore);
  });
});
