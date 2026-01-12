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

  it("tokenizes quoted sheet-qualified references with escaped apostrophes", () => {
    const tokens = tokenizeFormula("=SUM('Bob''s Sheet'!A1, 'Bob''s Sheet'!A1:A2)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["'Bob''s Sheet'!A1", "'Bob''s Sheet'!A1:A2"]);
  });

  it("tokenizes unquoted Unicode sheet-qualified references", () => {
    const tokens = tokenizeFormula("=SUM(résumé!A1, 数据!B2)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["résumé!A1", "数据!B2"]);
  });

  it("does not treat unquoted sheet names containing spaces as sheet-qualified refs", () => {
    const tokens = tokenizeFormula("=SUM(My Sheet!A1)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["A1"]);
  });

  it("tokenizes spill operator (#) and percent operator (%)", () => {
    const tokens = tokenizeFormula("=A1# + 1%").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["reference", "A1"],
      ["operator", "#"],
      ["operator", "+"],
      ["number", "1"],
      ["operator", "%"],
    ]);
  });

  it("does not include trailing brackets in error literals", () => {
    const tokens = tokenizeFormula("=[#REF!]").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["punctuation", "["],
      ["error", "#REF!"],
      ["punctuation", "]"],
    ]);
  });
});
