import { describe, expect, it } from "vitest";

import type { AIAuditEntry } from "../../../../../packages/ai-audit/src/types.js";

import { createAuditLogExport } from "./exportAuditLog";

describe("createAuditLogExport", () => {
  it("creates a JSON blob containing the provided audit entries", async () => {
    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1700000000000,
        session_id: "session-1",
        mode: "chat",
        input: { message: "hello" },
        model: "unit-test-model",
        latency_ms: 123,
        token_usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
        tool_calls: [{ name: "write_cell", parameters: { cell: "A1", value: 1 }, approved: true, ok: true }],
      },
    ];

    const exp = createAuditLogExport(entries, { fileName: "audit.json" });
    expect(exp.blob.type).toBe("application/json");
    expect(exp.fileName).toBe("audit.json");

    const text = await exp.blob.text();
    expect(JSON.parse(text)).toEqual(entries);
  });
});

