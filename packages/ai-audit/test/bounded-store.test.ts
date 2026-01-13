import { describe, expect, it } from "vitest";

import { BoundedAIAuditStore, MemoryAIAuditStore } from "../src/index.js";
import type { AIAuditEntry } from "../src/types.js";

describe("BoundedAIAuditStore", () => {
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
      input: { workbookId: "workbook-from-input", prompt: huge },
      model: "unit-test-model",
      tool_calls: [{ name: "tool", parameters: { payload: huge } }]
    });

    // Case 2: derive from session_id prefix when input has no workbook id.
    await store.logEntry({
      id: "entry-derive-2",
      timestamp_ms: Date.now(),
      session_id: "workbook-from-session:550e8400-e29b-41d4-a716-446655440000",
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
});
