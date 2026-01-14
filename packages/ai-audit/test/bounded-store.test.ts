import { describe, expect, it } from "vitest";

import { BoundedAIAuditStore, LocalStorageAIAuditStore, MemoryAIAuditStore } from "../src/index.js";
import type { AIAuditStore } from "../src/store.js";
import type { AIAuditEntry } from "../src/types.js";

describe("BoundedAIAuditStore", () => {
  it("defaults max_entry_chars to 200k and normalizes invalid values", () => {
    const storeA = new BoundedAIAuditStore(new MemoryAIAuditStore());
    expect(storeA.maxEntryChars).toBe(200_000);

    const storeB = new BoundedAIAuditStore(new MemoryAIAuditStore(), { max_entry_chars: -1 });
    expect(storeB.maxEntryChars).toBe(200_000);

    const storeC = new BoundedAIAuditStore(new MemoryAIAuditStore(), { max_entry_chars: Number.NaN });
    expect(storeC.maxEntryChars).toBe(200_000);
  });

  it("stores entries as-is when they are under the size cap", async () => {
    const underlying = new MemoryAIAuditStore();
    const maxEntryChars = 2_000;
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: maxEntryChars });

    const entry: AIAuditEntry = {
      id: "entry-small",
      timestamp_ms: Date.now(),
      session_id: "session-small",
      mode: "chat",
      input: { prompt: "hi" },
      model: "unit-test-model",
      tool_calls: [{ name: "tool", parameters: { a: 1 }, result: { ok: true } }]
    };

    expect(JSON.stringify(entry).length).toBeLessThanOrEqual(maxEntryChars);

    await store.logEntry(entry);

    const stored = await underlying.listEntries({ session_id: "session-small" });
    expect(stored.length).toBe(1);

    expect(stored[0]!.input).toEqual(entry.input);
    expect((stored[0]!.input as any)?.audit_truncated).toBeUndefined();
    expect(stored[0]!.tool_calls[0]!.result).toEqual({ ok: true });
  });

  it("does not compact entries solely due to BigInt values", async () => {
    const underlying = new MemoryAIAuditStore();
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: 2_000 });

    const entry: AIAuditEntry = {
      id: "entry-bigint",
      timestamp_ms: Date.now(),
      session_id: "session-bigint",
      mode: "chat",
      input: { big: 123n },
      model: "unit-test-model",
      tool_calls: [],
    };

    await store.logEntry(entry);

    const stored = await underlying.listEntries({ session_id: "session-bigint" });
    expect(stored).toHaveLength(1);
    expect((stored[0]!.input as any)?.audit_truncated).toBeUndefined();
    expect((stored[0]!.input as any).big).toBe(123n);
  });

  it("preserves verification fields containing BigInt when compacting oversized entries", async () => {
    const underlying = new MemoryAIAuditStore();
    const maxEntryChars = 2_000;
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: maxEntryChars });

    const huge = "x".repeat(50_000);
    const entry: AIAuditEntry = {
      id: "entry-bigint-verify",
      timestamp_ms: Date.now(),
      session_id: "session-bigint-verify",
      mode: "chat",
      input: { prompt: huge },
      model: "unit-test-model",
      tool_calls: [],
      verification: {
        needs_tools: false,
        used_tools: false,
        verified: true,
        confidence: 1,
        warnings: [],
        claims: [
          {
            claim: "bigint",
            verified: true,
            toolEvidence: { big: 123n },
          },
        ],
      },
    };

    // Sanity: entry cannot be JSON.stringified due to BigInt and is oversized due to input.
    expect(() => JSON.stringify(entry)).toThrow();

    await store.logEntry(entry);

    const stored = await underlying.listEntries({ session_id: "session-bigint-verify" });
    expect(stored).toHaveLength(1);

    const storedEntry = stored[0]!;
    expect(storedEntry.verification?.claims?.[0]?.toolEvidence).toEqual({ big: 123n });

    // Ensure the size cap is still respected when serializing BigInt values.
    const json = JSON.stringify(storedEntry, (_k, v) => (typeof v === "bigint" ? v.toString() : v));
    expect(json.length).toBeLessThanOrEqual(maxEntryChars);
  });

  it("compacts oversized entries to stay within the configured size limit", async () => {
    const underlying = new MemoryAIAuditStore();
    const maxEntryChars = 2_000;
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: maxEntryChars });

    const huge = "x".repeat(50_000);

    const entry: AIAuditEntry = {
      id: "entry-1",
      timestamp_ms: Date.now(),
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: huge },
      model: "unit-test-model",
      tool_calls: [
        {
          name: "huge_tool",
          parameters: { payload: huge },
          audit_result_summary: { summary: huge },
          result: { ok: true, payload: huge }
        }
      ]
    };

    // Sanity: the raw entry should exceed our configured limit.
    expect(JSON.stringify(entry).length).toBeGreaterThan(maxEntryChars);

    await store.logEntry(entry);

    const stored = await underlying.listEntries({ session_id: "session-1" });
    expect(stored.length).toBe(1);
    const storedEntry = stored[0]!;

    expect(JSON.stringify(storedEntry).length).toBeLessThanOrEqual(maxEntryChars);

    // Filter-critical fields are preserved.
    expect(storedEntry.id).toBe(entry.id);
    expect(storedEntry.timestamp_ms).toBe(entry.timestamp_ms);
    expect(storedEntry.session_id).toBe(entry.session_id);
    expect(storedEntry.workbook_id).toBe(entry.workbook_id);
    expect(storedEntry.mode).toBe(entry.mode);
    expect(storedEntry.model).toBe(entry.model);

    const toolCall = storedEntry.tool_calls[0]!;
    expect(toolCall.name).toBe("huge_tool");
    expect(toolCall.result).toBeUndefined();
    expect(toolCall.result_truncated).toBe(true);

    const input = storedEntry.input as any;
    expect(input?.audit_truncated).toBe(true);
    expect(typeof input?.audit_json).toBe("string");
    expect(typeof input?.audit_original_chars).toBe("number");
    expect(input.audit_original_chars).toBeGreaterThan((input.audit_json as string).length);

    const params = toolCall.parameters as any;
    expect(params?.audit_truncated).toBe(true);
    expect(typeof params?.audit_json).toBe("string");
    expect(typeof params?.audit_original_chars).toBe("number");
    expect(params.audit_original_chars).toBeGreaterThan((params.audit_json as string).length);

    const summary = toolCall.audit_result_summary as any;
    expect(summary?.audit_truncated).toBe(true);
    expect(typeof summary?.audit_json).toBe("string");
    expect(typeof summary?.audit_original_chars).toBe("number");
    expect(summary.audit_original_chars).toBeGreaterThan((summary.audit_json as string).length);

    // Compaction should not mutate the original entry object.
    expect((entry.tool_calls[0] as any).result).toEqual({ ok: true, payload: huge });
    expect((entry.input as any).prompt).toBe(huge);
  });

  it("derives workbook_id during compaction when missing/blank", async () => {
    const underlying = new MemoryAIAuditStore();
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: 2_000 });

    const huge = "x".repeat(50_000);

    // Case 1: derive from input.workbookId when workbook_id is missing/blank.
    await store.logEntry({
      id: "entry-derive-1",
      timestamp_ms: Date.now(),
      session_id: "session-derive-1",
      workbook_id: "   ",
      mode: "chat",
      input: { workbookId: "  workbook-from-input  ", prompt: huge },
      model: "unit-test-model",
      tool_calls: [{ name: "tool", parameters: { payload: huge } }]
    });

    // Case 2: derive from session_id prefix when input has no workbook id.
    await store.logEntry({
      id: "entry-derive-2",
      timestamp_ms: Date.now(),
      session_id: "  workbook-from-session  :550e8400-e29b-41d4-a716-446655440000",
      mode: "chat",
      input: { prompt: huge },
      model: "unit-test-model",
      tool_calls: [{ name: "tool", parameters: { payload: huge } }]
    });

    const entries = await underlying.listEntries();
    const byId = new Map(entries.map((e) => [e.id, e]));

    expect(byId.get("entry-derive-1")!.workbook_id).toBe("workbook-from-input");
    expect(byId.get("entry-derive-2")!.workbook_id).toBe("workbook-from-session");
  });

  it("prevents LocalStorageAIAuditStore persistence fallback by keeping entries under a quota", async () => {
    class QuotaLocalStorage implements Storage {
      #data = new Map<string, string>();
      constructor(private readonly maxChars: number) {}

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
        if (value.length > this.maxChars) {
          throw new Error("QuotaExceededError");
        }
        this.#data.set(key, value);
      }
    }

    const originalDescriptor = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
    Object.defineProperty(globalThis, "localStorage", { value: new QuotaLocalStorage(3_000), configurable: true });

    try {
      const huge = "x".repeat(50_000);
      const entry: AIAuditEntry = {
        id: "entry-quota-1",
        timestamp_ms: Date.now(),
        session_id: "session-quota-1",
        workbook_id: "workbook-quota-1",
        mode: "chat",
        input: { prompt: huge },
        model: "unit-test-model",
        tool_calls: [{ name: "tool", parameters: { payload: huge }, result: { ok: true, payload: huge } }]
      };

      // Without the bounded wrapper, LocalStorageAIAuditStore will hit the quota and
      // fall back to in-memory storage (no persisted localStorage value).
      const rawStore = new LocalStorageAIAuditStore({ key: "audit_quota_raw" });
      await rawStore.logEntry(entry);
      expect(globalThis.localStorage.getItem("audit_quota_raw")).toBeNull();

      // With the bounded wrapper, the entry is compacted so the underlying store can persist it.
      const bounded = new BoundedAIAuditStore(new LocalStorageAIAuditStore({ key: "audit_quota_bounded" }), {
        max_entry_chars: 1_500
      });
      await bounded.logEntry(entry);

      const persisted = globalThis.localStorage.getItem("audit_quota_bounded");
      expect(persisted).toBeTypeOf("string");
      expect((persisted as string).length).toBeLessThanOrEqual(3_000);

      const parsed = JSON.parse(persisted as string) as AIAuditEntry[];
      expect(parsed).toHaveLength(1);
      expect(parsed[0]!.tool_calls[0]!.result).toBeUndefined();
      expect((parsed[0]!.input as any)?.audit_truncated).toBe(true);
    } finally {
      if (originalDescriptor) {
        Object.defineProperty(globalThis, "localStorage", originalDescriptor);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      }
    }
  });

  it("delegates listEntries() to the underlying store", async () => {
    const underlying = new MemoryAIAuditStore();
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: 1_000 });

    const entry: AIAuditEntry = {
      id: "entry-list-1",
      timestamp_ms: 1,
      session_id: "session-list",
      mode: "chat",
      input: { prompt: "hi" },
      model: "unit-test-model",
      tool_calls: []
    };

    await underlying.logEntry(entry);
    const listed = await store.listEntries({ session_id: "session-list" });
    expect(listed.map((e) => e.id)).toEqual(["entry-list-1"]);
  });

  it("handles non-serializable (circular) entries by compacting instead of throwing", async () => {
    const underlying = new MemoryAIAuditStore();
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: 2_000 });

    const circular: any = {};
    circular.self = circular;

    const entry: AIAuditEntry = {
      id: "entry-circular-1",
      timestamp_ms: Date.now(),
      session_id: "session-circular-1",
      mode: "chat",
      input: circular,
      model: "unit-test-model",
      tool_calls: []
    };

    await expect(store.logEntry(entry)).resolves.toBeUndefined();

    const stored = await underlying.listEntries({ session_id: "session-circular-1" });
    expect(stored).toHaveLength(1);
    const storedInput = stored[0]!.input as any;
    expect(storedInput?.audit_truncated).toBe(true);
    expect(typeof storedInput?.audit_json).toBe("string");
  });

  it("drops non-serializable optional fields during compaction so underlying stores can JSON.stringify", async () => {
    class JsonStringifyStore implements AIAuditStore {
      readonly entries: AIAuditEntry[] = [];
      async logEntry(entry: AIAuditEntry): Promise<void> {
        JSON.stringify(entry);
        this.entries.push(entry);
      }
      async listEntries(): Promise<AIAuditEntry[]> {
        return this.entries.slice();
      }
    }

    const underlying = new JsonStringifyStore();
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: 2_000 });

    const circular: any = {};
    circular.self = circular;

    const huge = "x".repeat(50_000);

    const entry: AIAuditEntry = {
      id: "entry-circular-verification-1",
      timestamp_ms: Date.now(),
      session_id: "session-circular-verification-1",
      mode: "chat",
      input: { prompt: huge },
      model: "unit-test-model",
      tool_calls: [],
      verification: {
        needs_tools: false,
        used_tools: false,
        verified: false,
        confidence: 0,
        warnings: [],
        claims: [{ claim: "x", verified: false, toolEvidence: circular }]
      }
    };

    await expect(store.logEntry(entry)).resolves.toBeUndefined();

    const stored = underlying.entries[0]!;
    expect(stored.verification).toBeUndefined();
    expect((stored.input as any)?.audit_truncated).toBe(true);
    expect(JSON.stringify(stored).length).toBeLessThanOrEqual(2_000);
  });

  it("drops excess tool calls when needed to fit the entry budget", async () => {
    const underlying = new MemoryAIAuditStore();
    const maxEntryChars = 1_500;
    const store = new BoundedAIAuditStore(underlying, { max_entry_chars: maxEntryChars });

    const tool_calls = Array.from({ length: 50 }, (_, i) => ({ name: `tool_${i}`, parameters: { i } }));

    const entry: AIAuditEntry = {
      id: "entry-many-tools-1",
      timestamp_ms: Date.now(),
      session_id: "session-many-tools-1",
      mode: "chat",
      input: { prompt: "hello" },
      model: "unit-test-model",
      tool_calls,
    };

    expect(JSON.stringify(entry).length).toBeGreaterThan(maxEntryChars);

    await store.logEntry(entry);

    const stored = await underlying.listEntries({ session_id: "session-many-tools-1" });
    expect(stored).toHaveLength(1);

    const storedEntry = stored[0]!;
    expect(JSON.stringify(storedEntry).length).toBeLessThanOrEqual(maxEntryChars);
    expect(storedEntry.tool_calls.length).toBeLessThan(tool_calls.length);
    expect(storedEntry.tool_calls.some((c) => c.name === "audit_truncated_tool_calls")).toBe(true);
  });
});
