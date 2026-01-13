import { describe, expect, it } from "vitest";

import type { AIAuditEntry } from "../src/types.js";
import { serializeAuditEntries } from "../src/export.js";

describe("serializeAuditEntries", () => {
  it("formats NDJSON with one JSON object per line", () => {
    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input: { prompt: "hi" },
        model: "unit-test-model",
        tool_calls: []
      },
      {
        id: "audit-2",
        timestamp_ms: 2,
        session_id: "session",
        mode: "chat",
        input: { prompt: "bye" },
        model: "unit-test-model",
        tool_calls: []
      }
    ];

    const output = serializeAuditEntries(entries, { format: "ndjson", redactToolResults: false });
    const lines = output.split("\n");

    expect(lines).toHaveLength(2);
    expect(lines.map((line) => JSON.parse(line))).toEqual(entries);
  });

  it("redacts tool_calls[].result when enabled", () => {
    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input: { prompt: "hi" },
        model: "unit-test-model",
        tool_calls: [
          {
            name: "tool",
            parameters: { a: 1 },
            result: { secret: "should-not-export" },
            audit_result_summary: "ok"
          }
        ]
      }
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: true });
    const parsed = JSON.parse(output) as AIAuditEntry[];

    expect(parsed[0]!.tool_calls[0]).not.toHaveProperty("result");
    expect(parsed[0]!.tool_calls[0]!.audit_result_summary).toBe("ok");
  });

  it("truncates oversized audit_result_summary when enabled", () => {
    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input: { prompt: "hi" },
        model: "unit-test-model",
        tool_calls: [
          {
            name: "tool",
            parameters: { a: 1 },
            audit_result_summary: "x".repeat(50)
          }
        ]
      }
    ];

    const output = serializeAuditEntries(entries, {
      format: "json",
      redactToolResults: true,
      maxToolResultChars: 10
    });
    const parsed = JSON.parse(output) as Array<{
      tool_calls: Array<{ audit_result_summary?: unknown; export_truncated?: boolean }>;
    }>;

    const toolCall = parsed[0]!.tool_calls[0]!;
    expect(typeof toolCall.audit_result_summary).toBe("string");
    expect((toolCall.audit_result_summary as string).length).toBeLessThanOrEqual(10);
    expect(toolCall.export_truncated).toBe(true);
  });

  it("produces deterministic output regardless of object insertion order", () => {
    const entryA: AIAuditEntry = {
      id: "audit-1",
      timestamp_ms: 1,
      session_id: "session",
      mode: "chat",
      input: { b: 1, a: 2 },
      model: "unit-test-model",
      tool_calls: [
        {
          name: "tool",
          parameters: { z: 1, a: 2 },
          result: { b: 1, a: 2 }
        }
      ]
    };

    const entryB: AIAuditEntry = {
      id: "audit-1",
      timestamp_ms: 1,
      session_id: "session",
      mode: "chat",
      input: { a: 2, b: 1 },
      model: "unit-test-model",
      tool_calls: [
        {
          name: "tool",
          parameters: { a: 2, z: 1 },
          result: { a: 2, b: 1 }
        }
      ]
    };

    const outA = serializeAuditEntries([entryA], { format: "json", redactToolResults: false });
    const outB = serializeAuditEntries([entryB], { format: "json", redactToolResults: false });

    expect(outA).toBe(outB);
    expect(outA).toContain(`"input":{"a":2,"b":1}`);
  });
});

