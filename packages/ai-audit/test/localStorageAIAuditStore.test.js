import assert from "node:assert/strict";
import test from "node:test";

import { LocalStorageAIAuditStore } from "../src/local-storage-store.ts";
import { MemoryAIAuditStore } from "../src/memory-store.ts";

const originalLocalStorageDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
const originalStructuredCloneDescriptor = Object.getOwnPropertyDescriptor(globalThis, "structuredClone");
const originalDateNow = Date.now;

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

class MemoryLocalStorage {
  constructor() {
    this.data = new Map();
  }

  getItem(key) {
    return this.data.get(key) ?? null;
  }

  setItem(key, value) {
    this.data.set(key, value);
  }

  removeItem(key) {
    this.data.delete(key);
  }

  clear() {
    this.data.clear();
  }
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

test("LocalStorageAIAuditStore enforces max_age_ms on logEntry()", async () => {
  let now = 1_000_000;
  Date.now = () => now;

  try {
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    const key = "audit_test_max_age_log";
    const store = new LocalStorageAIAuditStore({ key, max_age_ms: 1_000 });

    await store.logEntry(makeEntry("old", "s1"));

    // Advance beyond max_age_ms then log a new entry; the old one should be purged at write-time.
    now += 1_500;
    await store.logEntry(makeEntry("new", "s1"));

    const raw = globalThis.localStorage.getItem(key);
    assert.ok(raw);
    const parsed = JSON.parse(raw);
    assert.deepEqual(
      parsed.map((entry) => entry.id),
      ["new"],
    );
  } finally {
    Date.now = originalDateNow;
    restoreLocalStorage();
  }
});

test("LocalStorageAIAuditStore enforces max_age_ms on listEntries() (best-effort purge)", async () => {
  let now = 2_000_000;
  Date.now = () => now;

  try {
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    const key = "audit_test_max_age_list";
    const store = new LocalStorageAIAuditStore({ key, max_age_ms: 1_000 });
    await store.logEntry(makeEntry("old", "s1"));

    // No new writes; listEntries() should still purge expired entries.
    now += 1_500;
    const entries = await store.listEntries({ session_id: "s1" });
    assert.equal(entries.length, 0);

    const raw = globalThis.localStorage.getItem(key);
    assert.ok(raw);
    assert.deepEqual(JSON.parse(raw), []);
  } finally {
    Date.now = originalDateNow;
    restoreLocalStorage();
  }
});

test("LocalStorageAIAuditStore enforces max_age_ms in memory fallback", async () => {
  let now = 3_000_000;
  Date.now = () => now;

  try {
    Object.defineProperty(globalThis, "localStorage", { value: undefined, configurable: true, writable: true });

    const store = new LocalStorageAIAuditStore({ key: "audit_test_max_age_memory", max_age_ms: 1_000 });
    await store.logEntry(makeEntry("old", "s1"));

    now += 1_500;
    assert.deepEqual(await store.listEntries({ session_id: "s1" }), []);

    // If listEntries() only filtered without persisting the purge, rewinding time could
    // make the old entry appear again.
    now -= 1_400;
    assert.deepEqual(await store.listEntries({ session_id: "s1" }), []);
  } finally {
    Date.now = originalDateNow;
    restoreLocalStorage();
  }
});
