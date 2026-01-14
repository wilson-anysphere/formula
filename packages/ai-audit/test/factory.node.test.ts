import { describe, expect, it, afterEach } from "vitest";

import { createDefaultAIAuditStore } from "../src/factory.node.js";
import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";

import { IDBKeyRange, indexedDB } from "fake-indexeddb";

const originalGlobals = {
  indexedDB: Object.getOwnPropertyDescriptor(globalThis as any, "indexedDB"),
  IDBKeyRange: Object.getOwnPropertyDescriptor(globalThis as any, "IDBKeyRange"),
};

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
});

describe("createDefaultAIAuditStore (node entrypoint)", () => {
  it("defaults to BoundedAIAuditStore wrapping MemoryAIAuditStore", async () => {
    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(MemoryAIAuditStore);
  });

  it("still defaults to memory even when indexedDB globals exist", async () => {
    Object.defineProperty(globalThis as any, "indexedDB", { value: indexedDB, configurable: true });
    Object.defineProperty(globalThis as any, "IDBKeyRange", { value: IDBKeyRange, configurable: true });

    const store = await createDefaultAIAuditStore();
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

