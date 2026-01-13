import { describe, expect, it } from "vitest";

import type { AIAuditEntry } from "../src/types.js";
import { serializeAuditEntries } from "../src/export.js";

describe("serializeAuditEntries", () => {
  it("defaults to NDJSON output with tool result redaction enabled", () => {
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
          },
        ],
      },
    ];

    const output = serializeAuditEntries(entries);
    const lines = output.split("\n");

    expect(lines).toHaveLength(1);
    const parsed = JSON.parse(lines[0]!) as AIAuditEntry;
    expect(Array.isArray(parsed)).toBe(false);
    expect(parsed.tool_calls[0]).not.toHaveProperty("result");
  });

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

  it("serializes BigInt values without throwing (deterministic)", () => {
    const big = 12345678901234567890n;

    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input: { big },
        model: "unit-test-model",
        tool_calls: [
          {
            name: "tool",
            parameters: { big },
            audit_result_summary: big,
          },
        ],
      },
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: true, maxToolResultChars: 10_000 });
    const parsed = JSON.parse(output) as any[];

    expect(parsed[0].input).toEqual({ big: big.toString() });
    expect(parsed[0].tool_calls[0].parameters).toEqual({ big: big.toString() });
    expect(parsed[0].tool_calls[0].audit_result_summary).toBe(big.toString());
  });

  it("replaces circular references with a stable placeholder", () => {
    const input: any = { a: 1 };
    input.self = input;

    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input,
        model: "unit-test-model",
        tool_calls: [],
      },
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: false });
    const parsed = JSON.parse(output) as any[];

    expect(parsed[0].input).toEqual({ a: 1, self: "[Circular]" });
  });

  it("does not throw when serializing objects with throwing getters", () => {
    const input: any = {};
    Object.defineProperty(input, "secret", {
      enumerable: true,
      get() {
        throw new Error("boom");
      },
    });

    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input,
        model: "unit-test-model",
        tool_calls: [],
      },
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: false });
    const parsed = JSON.parse(output) as any[];

    expect(parsed[0].input.secret).toBe("[Unserializable]");
  });

  it("preserves __proto__ keys without prototype pollution", () => {
    const input: any = { a: 1 };
    Object.defineProperty(input, "__proto__", {
      value: { polluted: true },
      enumerable: true,
      configurable: true,
      writable: true,
    });

    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input,
        model: "unit-test-model",
        tool_calls: [],
      },
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: false });
    const parsed = JSON.parse(output) as any[];

    expect(Object.getPrototypeOf(parsed[0].input)).toBe(Object.prototype);
    expect(parsed[0].input["__proto__"]).toEqual({ polluted: true });
    expect(({} as any).polluted).toBeUndefined();
  });

  it("handles toJSON() methods that return self (no infinite recursion)", () => {
    const input: any = { a: 1 };
    input.toJSON = function () {
      return this;
    };

    const entries: AIAuditEntry[] = [
      {
        id: "audit-1",
        timestamp_ms: 1,
        session_id: "session",
        mode: "chat",
        input,
        model: "unit-test-model",
        tool_calls: [],
      },
    ];

    const output = serializeAuditEntries(entries, { format: "json", redactToolResults: false });
    const parsed = JSON.parse(output) as any[];
    expect(parsed[0].input).toBe("[Circular]");
  });
});
