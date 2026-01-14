import { describe, expect, it, vi } from "vitest";

import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import type { AIAuditEntry } from "../src/types.js";

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

function makeEntry(params: { id: string; session_id: string; timestamp_ms: number }): AIAuditEntry {
  return {
    id: params.id,
    timestamp_ms: params.timestamp_ms,
    session_id: params.session_id,
    mode: "chat",
    input: { prompt: "hello" },
    model: "unit-test-model",
    tool_calls: [],
  };
}

describe("workbook_id filtering (legacy fallback)", () => {
  it("LocalStorageAIAuditStore matches legacy entries using input.workbookId", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    try {
      const store = new LocalStorageAIAuditStore({ key: "audit_test_workbook_filter" });
      const legacy: AIAuditEntry = {
        id: "legacy-entry-1",
        timestamp_ms: Date.now(),
        session_id: "session-legacy",
        mode: "chat",
        input: { workbookId: "  workbook-1  ", prompt: "hello" },
        model: "unit-test-model",
        tool_calls: [],
      };

      await store.logEntry(legacy);
      const matches = await store.listEntries({ workbook_id: "workbook-1" });
      expect(matches.map((e) => e.id)).toEqual(["legacy-entry-1"]);
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });

  it("MemoryAIAuditStore matches legacy entries using input.workbookId", async () => {
    const store = new MemoryAIAuditStore();
    const legacy: AIAuditEntry = {
      id: "legacy-entry-1",
      timestamp_ms: Date.now(),
      session_id: "session-legacy",
      mode: "chat",
      input: { workbookId: "  workbook-1  ", prompt: "hello" },
      model: "unit-test-model",
      tool_calls: [],
    };

    await store.logEntry(legacy);
    const matches = await store.listEntries({ workbook_id: "workbook-1" });
    expect(matches.map((e) => e.id)).toEqual(["legacy-entry-1"]);
  });
});

describe("LocalStorageAIAuditStore serialization", () => {
  it("persists entries containing BigInt values by exporting them as strings", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    try {
      const key = "audit_test_bigint";
      const store = new LocalStorageAIAuditStore({ key });

      const entry: AIAuditEntry = {
        id: "bigint-entry-1",
        timestamp_ms: Date.now(),
        session_id: "session-bigint",
        mode: "chat",
        input: { big: 123n },
        model: "unit-test-model",
        tool_calls: [],
      };

      await store.logEntry(entry);

      const raw = (globalThis as any).localStorage.getItem(key);
      expect(raw).toBeTruthy();
      const parsed = JSON.parse(raw!) as any[];
      expect(parsed[0]?.input?.big).toBe("123");

      const roundTrip = await store.listEntries({ session_id: "session-bigint" });
      expect((roundTrip[0]!.input as any).big).toBe("123");
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });
});

describe("LocalStorageAIAuditStore max_entries retention", () => {
  it("drops oldest entries deterministically when timestamps tie (id tiebreaker)", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    try {
      const store = new LocalStorageAIAuditStore({ key: "audit_test_max_entries_tiebreaker", max_entries: 2 });

      // Insert in an order that would previously make retention depend on insertion order.
      await store.logEntry(makeEntry({ id: "c", session_id: "s1", timestamp_ms: 1000 }));
      await store.logEntry(makeEntry({ id: "a", session_id: "s1", timestamp_ms: 1000 }));
      await store.logEntry(makeEntry({ id: "b", session_id: "s1", timestamp_ms: 1000 }));

      const entries = await store.listEntries({ session_id: "s1" });
      expect(entries.map((entry) => entry.id)).toEqual(["c", "b"]);
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });
});

describe("LocalStorageAIAuditStore age-based retention", () => {
  it("drops expired entries on logEntry()", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-01-01T00:00:00.000Z"));

    try {
      const key = "audit_test_age_retention_log";
      const store = new LocalStorageAIAuditStore({ key, max_age_ms: 1000 });

      await store.logEntry(makeEntry({ id: "old", session_id: "s1", timestamp_ms: Date.now() }));

      // Advance beyond max_age_ms and write a new entry. The previous entry should be purged on write.
      vi.setSystemTime(Date.now() + 1500);
      await store.logEntry(makeEntry({ id: "new", session_id: "s1", timestamp_ms: Date.now() }));

      const raw = (globalThis as any).localStorage.getItem(key);
      expect(raw).toBeTruthy();
      const parsed = JSON.parse(raw!) as AIAuditEntry[];
      expect(parsed.map((e) => e.id)).toEqual(["new"]);
    } finally {
      vi.useRealTimers();
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });

  it("drops expired entries on listEntries() (best-effort purge)", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-01-01T00:00:00.000Z"));

    try {
      const key = "audit_test_age_retention_list";
      const store = new LocalStorageAIAuditStore({ key, max_age_ms: 1000 });

      await store.logEntry(makeEntry({ id: "old", session_id: "s1", timestamp_ms: Date.now() }));

      // Advance beyond max_age_ms and only read. listEntries() should purge the expired entry.
      vi.setSystemTime(Date.now() + 1500);
      const results = await store.listEntries({ session_id: "s1" });
      expect(results).toEqual([]);

      const raw = (globalThis as any).localStorage.getItem(key);
      expect(raw).toBeTruthy();
      const parsed = JSON.parse(raw!) as AIAuditEntry[];
      expect(parsed).toEqual([]);
    } finally {
      vi.useRealTimers();
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });

  it("applies retention to the in-memory fallback when localStorage is unavailable", async () => {
    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: undefined, configurable: true, writable: true });

    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-01-01T00:00:00.000Z"));

    try {
      const store = new LocalStorageAIAuditStore({ key: "audit_test_age_retention_memory", max_age_ms: 1000 });

      await store.logEntry(makeEntry({ id: "old", session_id: "s1", timestamp_ms: Date.now() }));

      // Advance beyond max_age_ms then list; the entry should be purged from the in-memory store.
      vi.setSystemTime(Date.now() + 1500);
      expect(await store.listEntries({ session_id: "s1" })).toEqual([]);

      // If listEntries() only filtered without purging, rewinding time could make the entry "unexpired" again.
      vi.setSystemTime(new Date("2024-01-01T00:00:00.500Z"));
      expect(await store.listEntries({ session_id: "s1" })).toEqual([]);
    } finally {
      vi.useRealTimers();
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });
});
