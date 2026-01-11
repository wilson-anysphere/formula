import { describe, expect, it } from "vitest";

import { classifyQueryNeedsTools, verifyToolUsage } from "../src/llm/verification.js";

describe("classifyQueryNeedsTools", () => {
  it("returns true when attachments are present", () => {
    expect(classifyQueryNeedsTools({ userText: "hello", attachments: [{ type: "range" }] })).toBe(true);
  });

  it("returns true when A1-style references are present", () => {
    expect(classifyQueryNeedsTools({ userText: "What's the value in Sheet1!A1?", attachments: [] })).toBe(true);
    expect(classifyQueryNeedsTools({ userText: "sum a1:b2 please", attachments: [] })).toBe(true);
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
});
