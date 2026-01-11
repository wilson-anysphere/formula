// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { LocalStorageBinaryStorage } from "@formula/ai-audit/browser";
import type { AIAuditEntry } from "@formula/ai-audit/browser";

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

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

const require = createRequire(import.meta.url);
const wasmPath = require.resolve("sql.js/dist/sql-wasm.wasm");
const wasmDataUrl = `data:application/wasm;base64,${readFileSync(wasmPath).toString("base64")}`;

describe("SqliteAIAuditStore (jsdom + LocalStorageBinaryStorage)", () => {
  it("round-trips persisted audit entries", async () => {
    vi.resetModules();

    const originalNodeVersion = Object.getOwnPropertyDescriptor(process.versions, "node");
    const original = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new MemoryLocalStorage(), configurable: true });

    try {
      if (originalNodeVersion) {
        Object.defineProperty(process.versions, "node", { value: undefined, configurable: true });
      }

      const { SqliteAIAuditStore } = await import("@formula/ai-audit/sqlite");
      const storage = new LocalStorageBinaryStorage("ai_audit_db_test");
      const store = await SqliteAIAuditStore.create({
        storage,
        locateFile: (file: string) => (file.endsWith(".wasm") ? wasmDataUrl : file)
      });

       const entry: AIAuditEntry = {
         id: `entry_${Date.now()}`,
         timestamp_ms: Date.now(),
         session_id: "session-1",
         workbook_id: "workbook-1",
         mode: "chat",
         input: { prompt: "hello" },
         model: "unit-test-model",
         tool_calls: []
       };

      await store.logEntry(entry);

       const roundTrip = await SqliteAIAuditStore.create({
         storage,
         locateFile: (file: string) => (file.endsWith(".wasm") ? wasmDataUrl : file)
       });
       const entries = await roundTrip.listEntries({ session_id: "session-1", workbook_id: "workbook-1", mode: "chat" });

      expect(entries.length).toBe(1);
      expect(entries[0]!.id).toBe(entry.id);
      expect(entries[0]!.model).toBe("unit-test-model");
    } finally {
      if (original) Object.defineProperty(globalThis, "localStorage", original);
      if (originalNodeVersion) Object.defineProperty(process.versions, "node", originalNodeVersion);
      vi.resetModules();
    }
  }, 60_000);
});
