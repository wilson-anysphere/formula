import { describe, expect, it } from "vitest";

import type { AIAuditStore } from "../src/store.js";
import type { AIAuditEntry } from "../src/types.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import { InMemoryBinaryStorage } from "../src/storage.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";
import { locateSqlJsFileNode } from "../src/sqlite-store.node.js";

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

function makeEntry(id: string, timestamp_ms: number, mode: AIAuditEntry["mode"] = "chat"): AIAuditEntry {
  return {
    id,
    timestamp_ms,
    session_id: "session-1",
    mode,
    input: { prompt: id },
    model: "unit-test-model",
    tool_calls: []
  };
}

type StoreFactory = () => Promise<{ store: AIAuditStore; cleanup?: () => void | Promise<void> }>;

const STORES: Array<{ name: string; create: StoreFactory }> = [
  {
    name: "MemoryAIAuditStore",
    create: async () => ({ store: new MemoryAIAuditStore() })
  },
  {
    name: "LocalStorageAIAuditStore",
    create: async () => {
      const original = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
      Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });
      const store = new LocalStorageAIAuditStore({ key: `audit_test_filters_${Date.now()}_${Math.random()}` });
      return {
        store,
        cleanup: async () => {
          if (original) {
            Object.defineProperty(globalThis, "localStorage", original);
          } else {
            // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
            delete (globalThis as any).localStorage;
          }
        }
      };
    }
  },
  {
    name: "SqliteAIAuditStore",
    create: async () => {
      const storage = new InMemoryBinaryStorage();
      const store = await SqliteAIAuditStore.create({ storage, locateFile: locateSqlJsFileNode });
      return { store };
    }
  }
];

describe("AuditListFilters (time ranges + cursor pagination)", () => {
  for (const { name, create } of STORES) {
    describe(name, () => {
      it("filters by after_timestamp_ms (inclusive) and before_timestamp_ms (exclusive)", async () => {
        const { store, cleanup } = await create();
        try {
          await store.logEntry(makeEntry("t1", 1000));
          await store.logEntry(makeEntry("t2", 2000));
          await store.logEntry(makeEntry("t3", 3000));

          expect((await store.listEntries({ after_timestamp_ms: 2000 })).map((e) => e.id)).toEqual(["t3", "t2"]);
          expect((await store.listEntries({ after_timestamp_ms: 3000 })).map((e) => e.id)).toEqual(["t3"]);

          expect((await store.listEntries({ before_timestamp_ms: 3000 })).map((e) => e.id)).toEqual(["t2", "t1"]);
          expect((await store.listEntries({ before_timestamp_ms: 2000 })).map((e) => e.id)).toEqual(["t1"]);

          expect((await store.listEntries({ after_timestamp_ms: 1500, before_timestamp_ms: 3000 })).map((e) => e.id)).toEqual(["t2"]);
        } finally {
          await cleanup?.();
        }
      });

      it("paginates stably with cursor across identical timestamps", async () => {
        const { store, cleanup } = await create();
        try {
          await store.logEntry(makeEntry("a", 5000));
          await store.logEntry(makeEntry("b", 5000));
          await store.logEntry(makeEntry("c", 5000));
          await store.logEntry(makeEntry("z", 4000));

          const page1 = await store.listEntries({ limit: 2 });
          expect(page1.map((e) => e.id)).toEqual(["c", "b"]);

          const page2 = await store.listEntries({
            limit: 2,
            cursor: { before_timestamp_ms: page1[1]!.timestamp_ms, before_id: page1[1]!.id }
          });
          expect(page2.map((e) => e.id)).toEqual(["a", "z"]);

          expect([...page1, ...page2].map((e) => e.id)).toEqual(["c", "b", "a", "z"]);
        } finally {
          await cleanup?.();
        }
      });

      it("treats an empty mode array as no mode filter", async () => {
        const { store, cleanup } = await create();
        try {
          await store.logEntry(makeEntry("chat-1", 1000, "chat"));
          await store.logEntry(makeEntry("inline-1", 2000, "inline_edit"));

          const entries = await store.listEntries({ mode: [] });
          expect(entries.map((e) => e.id)).toEqual(["inline-1", "chat-1"]);
        } finally {
          await cleanup?.();
        }
      });
    });
  }
});
