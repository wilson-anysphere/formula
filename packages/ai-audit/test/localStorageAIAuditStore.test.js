import assert from "node:assert/strict";
import test from "node:test";

import { LocalStorageAIAuditStore } from "../src/local-storage-store.ts";

const originalLocalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");

function restoreLocalStorage() {
  if (originalLocalStorageDescriptor) {
    Object.defineProperty(globalThis, "localStorage", originalLocalStorageDescriptor);
  } else {
    // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
    delete globalThis.localStorage;
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

