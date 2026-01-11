import { describe, expect, it } from "vitest";

import { classifyQueryNeedsTools } from "../src/llm/verification.js";

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

