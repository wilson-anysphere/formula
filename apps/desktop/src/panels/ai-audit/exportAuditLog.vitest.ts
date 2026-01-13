import { describe, expect, it } from "vitest";

import type { AIAuditEntry } from "@formula/ai-audit/browser";

import { createAuditLogExport } from "./exportAuditLog";

describe("createAuditLogExport", () => {
  it("exports NDJSON by default (one JSON object per line) and redacts tool results", async () => {
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
        tool_calls: [
          {
            name: "write_cell",
            parameters: { cell: "A1", value: 1 },
            approved: true,
            ok: true,
            result: { secret: "should-not-export" },
          },
        ],
      },
      {
        id: "audit-2",
        timestamp_ms: 1700000001000,
        session_id: "session-1",
        mode: "chat",
        input: { message: "world" },
        model: "unit-test-model",
        tool_calls: [],
      },
    ];

    const exp = createAuditLogExport(entries);
    expect(exp.blob.type).toBe("application/x-ndjson");
    expect(exp.fileName).toMatch(/\.ndjson$/);

    const text = await exp.blob.text();
    const lines = text.split("\n");
    expect(lines).toHaveLength(entries.length);

    const parsed = lines.map((line) => JSON.parse(line));
    expect(parsed.map((entry: AIAuditEntry) => entry.id)).toEqual(["audit-1", "audit-2"]);

    // Tool call result payloads should be removed by default.
    expect(parsed[0]?.tool_calls?.[0]).not.toHaveProperty("result");
  });

  it("exports a JSON array when format: json is requested", async () => {
    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1700000000000,
        session_id: "session-1",
        mode: "chat",
        input: { message: "hello" },
        model: "unit-test-model",
        tool_calls: [
          {
            name: "write_cell",
            parameters: { cell: "A1", value: 1 },
            approved: true,
            ok: true,
            result: { secret: "should-not-export" },
          },
        ],
      },
    ];

    const exp = createAuditLogExport(entries, { format: "json", fileName: "audit.json" });
    expect(exp.blob.type).toBe("application/json");
    expect(exp.fileName).toBe("audit.json");

    const text = await exp.blob.text();
    const parsed = JSON.parse(text) as AIAuditEntry[];
    expect(parsed).toHaveLength(1);
    expect(parsed[0]?.id).toBe("audit-1");
    expect(parsed[0]?.tool_calls?.[0]).not.toHaveProperty("result");
  });
});
