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

  it("tokenizes unquoted non-BMP Unicode sheet-qualified references", () => {
    const tokens = tokenizeFormula("=SUM(𝔘!A1, 𐐷!B2)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["𝔘!A1", "𐐷!B2"]);

    const input = "=SUM(𝔘!A1, 𐐷!B2)";
    const first = tokens.find((t) => t.type === "reference" && t.text === "𝔘!A1");
    const second = tokens.find((t) => t.type === "reference" && t.text === "𐐷!B2");
    expect(first).toBeTruthy();
    expect(second).toBeTruthy();
    expect(first?.start).toBe(input.indexOf("𝔘!A1"));
    expect(first?.end).toBe(input.indexOf("𝔘!A1") + "𝔘!A1".length);
    expect(second?.start).toBe(input.indexOf("𐐷!B2"));
    expect(second?.end).toBe(input.indexOf("𐐷!B2") + "𐐷!B2".length);
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

  it("tokenizes Excel structured references as single reference tokens", () => {
    const input = "=SUM(Table1[Amount])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference");
    expect(refs.map((t) => t.text)).toEqual(["Table1[Amount]"]);
    expect(refs[0]).toMatchObject({ start: input.indexOf("Table1"), end: input.indexOf("Table1") + "Table1[Amount]".length });
  });

  it("tokenizes structured references with nested brackets (e.g. #All) as single tokens", () => {
    const input = "=SUM(Table1[[#All],[Amount]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference");
    expect(refs.map((t) => t.text)).toEqual(["Table1[[#All],[Amount]]"]);
    expect(refs[0]).toMatchObject({
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount]]".length,
    });
  });

  it("tokenizes structured references with selectors (#Headers/#Data) as single tokens", () => {
    const headers = "=SUM(Table1[[#Headers],[Amount]])";
    const headerTokens = tokenizeFormula(headers);
    const headerRefs = headerTokens.filter((t) => t.type === "reference");
    expect(headerRefs.map((t) => t.text)).toEqual(["Table1[[#Headers],[Amount]]"]);

    const data = "=SUM(Table1[[#Data],[Amount]])";
    const dataTokens = tokenizeFormula(data);
    const dataRefs = dataTokens.filter((t) => t.type === "reference");
    expect(dataRefs.map((t) => t.text)).toEqual(["Table1[[#Data],[Amount]]"]);

    const totals = "=SUM(Table1[[#Totals],[Amount]])";
    const totalsTokens = tokenizeFormula(totals);
    const totalsRefs = totalsTokens.filter((t) => t.type === "reference");
    expect(totalsRefs.map((t) => t.text)).toEqual(["Table1[[#Totals],[Amount]]"]);
  });
});
