import { describe, expect, it } from "vitest";

import { tokenizeFormula } from "./tokenizeFormula.js";

describe("tokenizeFormula", () => {
  it("highlights functions, refs, strings, numbers, operators", () => {
    const tokens = tokenizeFormula('=IF(SUM(A1:A2) > 1000, "Over", "Within")');

    const types = tokens.filter((t) => t.type !== "whitespace").map((t) => [t.type, t.text]);

    expect(types.slice(0, 6)).toEqual([
      ["operator", "="],
      ["function", "IF"],
      ["punctuation", "("],
      ["function", "SUM"],
      ["punctuation", "("],
      ["reference", "A1:A2"],
    ]);

    expect(types.some(([type, text]) => type === "number" && text === "1000")).toBe(true);
    expect(types.some(([type]) => type === "operator")).toBe(true);
    expect(types.some(([type, text]) => type === "string" && text === '"Over"')).toBe(true);
  });
});
