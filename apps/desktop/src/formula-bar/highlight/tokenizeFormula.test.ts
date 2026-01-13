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

  it("does not treat ambiguous unquoted sheet prefixes as sheet-qualified refs", () => {
    const tokens = tokenizeFormula("=SUM(TRUE!A1, A1!B2, R1C1!C3)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["A1", "A1", "B2", "C3"]);
  });

  it("does not treat identifiers starting with cell-ref prefixes as references", () => {
    const tokens = tokenizeFormula("=A1FOO + R1C1FOO + A1.Price");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["A1"]);

    const idents = tokens.filter((t) => t.type === "identifier").map((t) => t.text);
    expect(idents).toEqual(expect.arrayContaining(["A1FOO", "R1C1FOO", "Price"]));
  });

  it("tokenizes unquoted external workbook and 3D sheet-qualified references", () => {
    const tokens = tokenizeFormula("=SUM([Book.xlsx]Sheet1!A1, Sheet1:Sheet3!B2)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[Book.xlsx]Sheet1!A1", "Sheet1:Sheet3!B2"]);
  });

  it("tokenizes Excel structured table references as a single reference token", () => {
    const tokens = tokenizeFormula("=SUM(Table1[Amount])");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[Amount]"]);
  });

  it("tokenizes structured refs with #All (including internal commas) as a single reference token", () => {
    const tokens = tokenizeFormula("=SUM(Table1[[#All],[Amount]])");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[[#All],[Amount]]"]);
  });
});
