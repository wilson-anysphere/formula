import { describe, expect, it } from "vitest";

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
        input: { workbookId: "workbook-1", prompt: "hello" },
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
      input: { workbookId: "workbook-1", prompt: "hello" },
      model: "unit-test-model",
      tool_calls: [],
    };

    await store.logEntry(legacy);
    const matches = await store.listEntries({ workbook_id: "workbook-1" });
    expect(matches.map((e) => e.id)).toEqual(["legacy-entry-1"]);
  });
});

