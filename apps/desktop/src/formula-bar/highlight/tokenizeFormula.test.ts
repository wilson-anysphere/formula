import { describe, expect, it } from "vitest";

import { tokenizeFormula } from "@formula/spreadsheet-frontend/formula/tokenizeFormula";

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

  it("treats whitespace between function names and '(' as a function token", () => {
    const tokens = tokenizeFormula("=SUM (A1, B1)").filter((t) => t.type !== "whitespace");
    expect(tokens.slice(0, 3).map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["function", "SUM"],
      ["punctuation", "("],
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

  it("tokenizes localized error literals with inverted punctuation (es-ES #¡VALOR!)", () => {
    const tokens = tokenizeFormula("=#¡VALOR! + 1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["error", "#¡VALOR!"],
      ["operator", "+"],
      ["number", "1"],
    ]);
  });

  it("tokenizes localized error literals with inverted question marks (es-ES #¿NOMBRE?)", () => {
    const tokens = tokenizeFormula("=#¿NOMBRE? + 1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["error", "#¿NOMBRE?"],
      ["operator", "+"],
      ["number", "1"],
    ]);
  });

  it("tokenizes localized error literals with short localized names (fr-FR #NOM?)", () => {
    const tokens = tokenizeFormula("=#NOM? + 1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["error", "#NOM?"],
      ["operator", "+"],
      ["number", "1"],
    ]);
  });

  it("tokenizes localized error literals with non-ASCII letters (de-DE #ÜBERLAUF!)", () => {
    const tokens = tokenizeFormula("=#ÜBERLAUF! + 1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["error", "#ÜBERLAUF!"],
      ["operator", "+"],
      ["number", "1"],
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

  it("tokenizes external workbook references with escaped closing brackets in the workbook name", () => {
    // Excel encodes literal `]` characters inside the `[Book]` segment by doubling them: `]]`.
    const tokens = tokenizeFormula("=SUM([Book]]Name.xlsx]Sheet1!A1, 1)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[Book]]Name.xlsx]Sheet1!A1"]);
  });

  it("tokenizes external workbook references whose workbook name contains '[' characters (non-nesting)", () => {
    // Workbook names may contain literal `[` without introducing nesting. The workbook prefix ends
    // at the first non-escaped `]` before the sheet name.
    const tokens = tokenizeFormula("=SUM([A1[Name.xlsx]Sheet1!A1, 1)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[A1[Name.xlsx]Sheet1!A1"]);
  });

  it("tokenizes external workbook references with quoted sheet names after an unquoted workbook prefix", () => {
    // Excel permits quoting the sheet token even when the workbook prefix itself is unquoted.
    // The engine serializer will typically quote the whole sheet spec instead, but we still
    // tokenize this form for best-effort highlighting parity.
    const tokens = tokenizeFormula("=SUM([Book.xlsx]'My Sheet'!A1, 1)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[Book.xlsx]'My Sheet'!A1"]);
  });

  it("tokenizes external workbook 3D sheet spans with quoted sheet tokens after an unquoted workbook prefix", () => {
    const tokens = tokenizeFormula("=SUM([Book.xlsx]'Sheet 1':'Sheet 3'!A1, 1)");
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[Book.xlsx]'Sheet 1':'Sheet 3'!A1"]);
  });

  it("tokenizes workbook-scoped external defined names (quoted external name refs)", () => {
    // The engine serializer emits workbook-scoped external defined names as a single quoted token:
    //   `'[Book.xlsx]MyName'`
    // Tokenize it as an identifier so syntax highlighting doesn't fall back to "unknown" tokens.
    const tokens = tokenizeFormula("='[Book.xlsx]MyName'+1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["identifier", "'[Book.xlsx]MyName'"],
      ["operator", "+"],
      ["number", "1"],
    ]);
  });

  it("tokenizes workbook-scoped external defined names (unquoted external name refs)", () => {
    // The engine parser accepts unquoted workbook-scoped external names (and users may type them)
    // even though the canonical renderer prefers the quoted form.
    const tokens = tokenizeFormula("=[Book.xlsx]MyName+1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["identifier", "[Book.xlsx]MyName"],
      ["operator", "+"],
      ["number", "1"],
    ]);
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

  it("tokenizes multi-column structured references as a single reference token", () => {
    const input = "=SUM(Table1[[#All],[Col1],[Col2]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[[#All],[Col1],[Col2]]"]);
    // Ensure `#All` is not mis-tokenized as an error literal.
    expect(tokens.filter((t) => t.type === "error").map((t) => t.text)).toEqual([]);
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

  it("tokenizes structured references with @ shorthand (this row) as single tokens", () => {
    const input = "=SUM(Table1[@Amount], Table1[@[Total Amount]], Table1[@])";
    const refs = tokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    expect(refs).toEqual(["Table1[@Amount]", "Table1[@[Total Amount]]", "Table1[@]"]);
  });

  it("tokenizes structured references with `#This Row` selectors as single tokens", () => {
    const input = "=SUM(Table1[[#This Row],[Amount]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[[#This Row],[Amount]]"]);
  });

  it("tokenizes structured references with @Column shorthand as single tokens", () => {
    const input = "=SUM(Table1[@Amount])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[@Amount]"]);
  });

  it("tokenizes structured references with @[[Column Name]] shorthand as single tokens", () => {
    const input = "=SUM(Table1[@[Total Amount]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["Table1[@[Total Amount]]"]);
  });

  it("tokenizes implicit this-row structured references ([@Column]) as single tokens", () => {
    const input = "=SUM([@Amount], 1)";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[@Amount]"]);
  });

  it("tokenizes implicit this-row structured references with nested bracket column names ([@[Column Name]]) as single tokens", () => {
    const input = "=SUM([@[Total Amount]], 1)";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference").map((t) => t.text);
    expect(refs).toEqual(["[@[Total Amount]]"]);
  });

  it("tokenizes structured references followed by operators (no trailing parens/commas)", () => {
    // Regression: bracket escaping logic should not prevent recognizing structured refs when they
    // are followed by an operator (e.g. `...]]+1`).
    const tokens = tokenizeFormula("=Table1[[#All],[Amount]]+1").filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["reference", "Table1[[#All],[Amount]]"],
      ["operator", "+"],
      ["number", "1"],
    ]);
  });

  it("does not consume into string literals when disambiguating `]]` in structured references", () => {
    // Regression: if a formula contains a structured ref followed by a string literal that
    // includes `]]`, we should not extend the structured ref token into the string.
    const input = '=SUM(Table1[[#All],[Amount]] & "]]", 1)';
    const tokens = tokenizeFormula(input).filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["function", "SUM"],
      ["punctuation", "("],
      ["reference", "Table1[[#All],[Amount]]"],
      ["operator", "&"],
      ["string", '"]]"'],
      ["punctuation", ","],
      ["number", "1"],
      ["punctuation", ")"],
    ]);
  });

  it("does not consume into string literals when tokenizing implicit structured references", () => {
    const input = '=SUM([[#All],[Amount]] & "]]", 1)';
    const tokens = tokenizeFormula(input).filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["function", "SUM"],
      ["punctuation", "("],
      ["reference", "[[#All],[Amount]]"],
      ["operator", "&"],
      ["string", '"]]"'],
      ["punctuation", ","],
      ["number", "1"],
      ["punctuation", ")"],
    ]);
  });

  it("does not consume into string literals when tokenizing nested implicit this-row references ([@[Column]])", () => {
    const input = '=SUM([@[Total Amount]] & "]]", 1)';
    const tokens = tokenizeFormula(input).filter((t) => t.type !== "whitespace");
    expect(tokens.map((t) => [t.type, t.text])).toEqual([
      ["operator", "="],
      ["function", "SUM"],
      ["punctuation", "("],
      ["reference", "[@[Total Amount]]"],
      ["operator", "&"],
      ["string", '"]]"'],
      ["punctuation", ","],
      ["number", "1"],
      ["punctuation", ")"],
    ]);
  });

  it("tokenizes structured table specifiers (#All/#Headers/#Data/#Totals) as single reference tokens", () => {
    const input = "=SUM(Table1[#All], Table1[#Headers], Table1[#Data], Table1[#Totals])";
    const refs = tokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    expect(refs).toEqual(["Table1[#All]", "Table1[#Headers]", "Table1[#Data]", "Table1[#Totals]"]);
  });

  it("tokenizes structured references with escaped closing brackets in column names as single tokens", () => {
    const input = "=COUNTA(Table1[[#Headers],[A]]B]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference");
    expect(refs.map((t) => t.text)).toEqual(["Table1[[#Headers],[A]]B]]"]);
    expect(refs[0]).toMatchObject({
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#Headers],[A]]B]]".length,
    });
  });

  it("tokenizes structured references where escaped `]` is followed by operator characters inside the column name", () => {
    const input = "=COUNTA(Table1[[#Headers],[A]]+B]])";
    const tokens = tokenizeFormula(input);
    const refs = tokens.filter((t) => t.type === "reference");
    expect(refs.map((t) => t.text)).toEqual(["Table1[[#Headers],[A]]+B]]"]);
  });
});
