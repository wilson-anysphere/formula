import { describe, expect, it } from "vitest";

import { classifyQueryNeedsTools, verifyAssistantClaims, verifyToolUsage } from "../src/llm/verification.js";

describe("classifyQueryNeedsTools", () => {
  it("returns true when attachments are present", () => {
    expect(classifyQueryNeedsTools({ userText: "hello", attachments: [{ type: "range" }] })).toBe(true);
  });

  it("returns true when A1-style references are present", () => {
    expect(classifyQueryNeedsTools({ userText: "What's the value in Sheet1!A1?", attachments: [] })).toBe(true);
    expect(classifyQueryNeedsTools({ userText: "sum a1:b2 please", attachments: [] })).toBe(true);
  });

  it("returns true when absolute/mixed A1-style references are present", () => {
    expect(classifyQueryNeedsTools({ userText: "What's in $A$1?" })).toBe(true);
    expect(classifyQueryNeedsTools({ userText: "sum A$1:A$3" })).toBe(true);
  });

  it("returns true when spreadsheet keywords are present", () => {
    expect(classifyQueryNeedsTools({ userText: "Compute the average of column B", attachments: [] })).toBe(true);
    expect(classifyQueryNeedsTools({ userText: "Create a pivot table", attachments: [] })).toBe(true);
  });

  it("returns false for generic questions without spreadsheet context", () => {
    expect(classifyQueryNeedsTools({ userText: "Tell me a joke", attachments: [] })).toBe(false);
  });
});

describe("verifyToolUsage", () => {
  it("treats successful tool execution as verified when tools were needed (e.g. write actions)", () => {
    const result = verifyToolUsage({
      needsTools: true,
      toolCalls: [{ name: "write_cell", ok: true }]
    });

    expect(result.needs_tools).toBe(true);
    expect(result.used_tools).toBe(true);
    expect(result.verified).toBe(true);
    expect(result.confidence).toBeGreaterThan(0.5);
  });

  it("requires a verified read/compute tool when requiredToolKind is 'verified'", () => {
    const result = verifyToolUsage({
      needsTools: true,
      requiredToolKind: "verified",
      toolCalls: [{ name: "write_cell", ok: true }]
    });

    expect(result.needs_tools).toBe(true);
    expect(result.used_tools).toBe(true);
    expect(result.verified).toBe(false);
  });
});

describe("verifyAssistantClaims", () => {
  it("treats count claims as exact matches (no floating tolerance)", async () => {
    const toolExecutor = {
      tools: [{ name: "compute_statistics" }],
      async execute(call: any) {
        return {
          tool: "compute_statistics",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments?.range, statistics: { count: 1_000_000 } }
        };
      }
    };

    const result = await verifyAssistantClaims({
      assistantText: "Count for Sheet1!A1:A10 is 1000001.",
      toolExecutor: toolExecutor as any
    });

    expect(result).not.toBeNull();
    expect(result?.verified).toBe(false);
    expect(result?.claims?.[0]).toMatchObject({
      verified: false,
      expected: 1000001,
      actual: 1000000
    });

    const evidence = (result?.claims?.[0] as any)?.toolEvidence;
    expect(evidence?.call?.name).toBe("compute_statistics");
    expect(evidence?.result?.data?.statistics?.count).toBe(1_000_000);
  });

  it("verifies cell_value claims when read_range returns formatted numeric strings (thousands separators)", async () => {
    const toolExecutor = {
      tools: [{ name: "read_range" }],
      async execute(call: any) {
        return {
          tool: "read_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments?.range, values: [["1,200"]] }
        };
      }
    };

    const result = await verifyAssistantClaims({
      assistantText: "Sheet1!A1 is 1200.",
      toolExecutor: toolExecutor as any
    });

    expect(result).not.toBeNull();
    expect(result?.verified).toBe(true);
    expect(result?.claims?.[0]).toMatchObject({
      verified: true,
      expected: 1200,
      actual: 1200
    });
  });

  it("verifies cell_value claims when read_range returns percent-formatted numeric strings", async () => {
    const toolExecutor = {
      tools: [{ name: "read_range" }],
      async execute(call: any) {
        return {
          tool: "read_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments?.range, values: [["10%"]] }
        };
      }
    };

    const result = await verifyAssistantClaims({
      assistantText: "Sheet1!A1 is 0.1.",
      toolExecutor: toolExecutor as any
    });

    expect(result).not.toBeNull();
    expect(result?.verified).toBe(true);
    expect(result?.claims?.[0]).toMatchObject({
      verified: true,
      expected: 0.1,
      actual: 0.1
    });
  });
});
