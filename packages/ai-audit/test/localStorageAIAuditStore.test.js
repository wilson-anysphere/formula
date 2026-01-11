import assert from "node:assert/strict";
import test from "node:test";

import { LocalStorageAIAuditStore } from "../src/local-storage-store.ts";
import { MemoryAIAuditStore } from "../src/memory-store.ts";

const originalLocalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalStructuredCloneDescriptor = Object.getOwnPropertyDescriptor(globalThis, "structuredClone");

function restoreLocalStorage() {
  if (originalLocalStorageDescriptor) {
    Object.defineProperty(globalThis, "localStorage", originalLocalStorageDescriptor);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete globalThis.localStorage;
  }
}

function restoreStructuredClone() {
  if (originalStructuredCloneDescriptor) {
    Object.defineProperty(globalThis, "structuredClone", originalStructuredCloneDescriptor);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete globalThis.structuredClone;
  }
}

function makeEntry(id, session_id) {
  return {
    id,
    timestamp_ms: Date.now(),
    session_id,
    mode: "chat",
    input: { text: "hi" },
    model: "test-model",
    tool_calls: [],
  };
}

test("LocalStorageAIAuditStore falls back to memory when localStorage is missing", async () => {
  try {
    Object.defineProperty(globalThis, "localStorage", { value: undefined, configurable: true, writable: true });

    const store = new LocalStorageAIAuditStore({ key: "audit_test_missing" });
    await store.logEntry(makeEntry("1", "s1"));
    const entries = await store.listEntries({ session_id: "s1" });
    assert.equal(entries.length, 1);
    assert.equal(entries[0].id, "1");
  } finally {
    restoreLocalStorage();
  }
});

test("LocalStorageAIAuditStore falls back to memory when localStorage getter throws", async () => {
  try {
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      get() {
        throw new Error("no localStorage");
      },
    });

    const store = new LocalStorageAIAuditStore({ key: "audit_test_throw_get" });
    await store.logEntry(makeEntry("2", "s2"));
    const entries = await store.listEntries({ session_id: "s2" });
    assert.equal(entries.length, 1);
    assert.equal(entries[0].id, "2");
  } finally {
    restoreLocalStorage();
  }
});

test("LocalStorageAIAuditStore sticks to memory fallback when setItem fails", async () => {
  try {
    const memory = new Map();
    Object.defineProperty(globalThis, "localStorage", {
      configurable: true,
      value: {
        getItem(key) {
          return memory.get(key) ?? null;
        },
        setItem() {
          throw new Error("quota exceeded");
        },
      },
    });

    const store = new LocalStorageAIAuditStore({ key: "audit_test_throw_set" });
    await store.logEntry(makeEntry("3", "s3"));

    // setItem throws so persistence fails, but the store should still behave consistently
    // by switching to in-memory storage for subsequent reads.
    const entries = await store.listEntries({ session_id: "s3" });
    assert.equal(entries.length, 1);
    assert.equal(entries[0].id, "3");
  } finally {
    restoreLocalStorage();
  }
});

test("Audit stores fall back to JSON cloning when structuredClone is missing", async () => {
  try {
    Object.defineProperty(globalThis, "structuredClone", { value: undefined, configurable: true, writable: true });
    Object.defineProperty(globalThis, "localStorage", { value: undefined, configurable: true, writable: true });

    const localStore = new LocalStorageAIAuditStore({ key: "audit_test_no_structured_clone" });
    await localStore.logEntry(makeEntry("4", "s4"));
    const localEntries = await localStore.listEntries({ session_id: "s4" });
    assert.equal(localEntries.length, 1);
    assert.equal(localEntries[0].id, "4");

    const memoryStore = new MemoryAIAuditStore();
    await memoryStore.logEntry(makeEntry("5", "s5"));
    const memoryEntries = await memoryStore.listEntries({ session_id: "s5" });
    assert.equal(memoryEntries.length, 1);
    assert.equal(memoryEntries[0].id, "5");
  } finally {
    restoreLocalStorage();
    restoreStructuredClone();
  }
});
