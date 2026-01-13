import { describe, expect, it } from "vitest";

import { BoundedAIAuditStore, MemoryAIAuditStore } from "../src/index.js";
import type { AIAuditEntry } from "../src/types.js";

describe("BoundedAIAuditStore", () => {
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

    expect(storedEntry.tool_calls[0]!.result).toBeUndefined();

    const input = storedEntry.input as any;
    expect(input?.audit_truncated).toBe(true);
  });
});

