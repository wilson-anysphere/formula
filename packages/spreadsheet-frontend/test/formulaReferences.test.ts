import { describe, expect, it } from "vitest";

import { assignFormulaReferenceColors, extractFormulaReferences, FORMULA_REFERENCE_PALETTE } from "../src/formulaReferences";

describe("extractFormulaReferences", () => {
  it("extracts simple A1 references with stable indices", () => {
    const { references, activeIndex } = extractFormulaReferences("=A1+B1", 0, 0);
    expect(activeIndex).toBe(null);
    expect(references).toEqual([
      {
        text: "A1",
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, sheet: undefined },
        index: 0,
        start: 1,
        end: 3
      },
      {
        text: "B1",
        range: { startRow: 0, startCol: 1, endRow: 0, endCol: 1, sheet: undefined },
        index: 1,
        start: 4,
        end: 6
      }
    ]);
  });

  it("parses sheet-qualified ranges", () => {
    const { references } = extractFormulaReferences("=SUM('My Sheet'!$A$1:$B$2)", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("'My Sheet'!$A$1:$B$2");
    expect(references[0]?.range).toEqual({ sheet: "My Sheet", startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  });

  it("parses sheet-qualified refs with escaped apostrophes", () => {
    const { references } = extractFormulaReferences("=SUM('O''Brien'!A1)", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("'O''Brien'!A1");
    expect(references[0]?.range).toEqual({ sheet: "O'Brien", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });

  it("parses unquoted Unicode sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=résumé!A1+数据!B2", 0, 0);
    expect(references).toHaveLength(2);
    expect(references[0]?.text).toBe("résumé!A1");
    expect(references[0]?.range).toEqual({ sheet: "résumé", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(references[1]?.text).toBe("数据!B2");
    expect(references[1]?.range).toEqual({ sheet: "数据", startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
  });

  it("parses unquoted non-BMP Unicode sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=𝔘!A1+𐐷!B2", 0, 0);
    expect(references).toHaveLength(2);
    expect(references[0]?.text).toBe("𝔘!A1");
    expect(references[0]?.range).toEqual({ sheet: "𝔘", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(references[1]?.text).toBe("𐐷!B2");
    expect(references[1]?.range).toEqual({ sheet: "𐐷", startRow: 1, startCol: 1, endRow: 1, endCol: 1 });

    // Ensure offsets are code-unit based (matches DOM selectionStart/selectionEnd semantics).
    const input = "=𝔘!A1+𐐷!B2";
    expect(references[0]?.start).toBe(input.indexOf("𝔘!A1"));
    expect(references[0]?.end).toBe(input.indexOf("𝔘!A1") + "𝔘!A1".length);
    expect(references[1]?.start).toBe(input.indexOf("𐐷!B2"));
    expect(references[1]?.end).toBe(input.indexOf("𐐷!B2") + "𐐷!B2".length);
  });

  it("does not treat invalid unquoted sheet names with spaces as sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=My Sheet!A1", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("A1");
    expect(references[0]?.range).toEqual({ sheet: undefined, startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });

  it("does not treat ambiguous unquoted sheet prefixes as sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=TRUE!A1 + A1!B2 + R1C1!C3", 0, 0);
    expect(references.map((r) => r.text)).toEqual(["A1", "A1", "B2", "C3"]);
    expect(references.map((r) => r.range.sheet)).toEqual([undefined, undefined, undefined, undefined]);
  });

  it("does not treat identifiers starting with cell-ref prefixes as references", () => {
    const { references } = extractFormulaReferences("=A1FOO + R1C1FOO + A1.Price", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("A1");
  });

  it("parses external workbook and 3D sheet-qualified references", () => {
    const { references } = extractFormulaReferences("=[Book.xlsx]Sheet1!A1 + Sheet1:Sheet3!B2", 0, 0);
    expect(references).toHaveLength(2);
    expect(references[0]?.text).toBe("[Book.xlsx]Sheet1!A1");
    expect(references[0]?.range).toEqual({ sheet: "[Book.xlsx]Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    expect(references[1]?.text).toBe("Sheet1:Sheet3!B2");
    expect(references[1]?.range).toEqual({ sheet: "Sheet1:Sheet3", startRow: 1, startCol: 1, endRow: 1, endCol: 1 });
  });

  it("parses 3D sheet-qualified refs with individually quoted sheet tokens", () => {
    const { references } = extractFormulaReferences("=SUM('Sheet 1':'Sheet 3'!A1, 1)", 0, 0);
    expect(references).toHaveLength(1);
    expect(references[0]?.text).toBe("'Sheet 1':'Sheet 3'!A1");
    expect(references[0]?.range).toEqual({ sheet: "Sheet 1:Sheet 3", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });

  it("detects the active reference at the caret (including token end)", () => {
    // =A1+B1, caret after final "1" should count as being in B1.
    const input = "=A1+B1";
    const { activeIndex } = extractFormulaReferences(input, input.length, input.length);
    expect(activeIndex).toBe(1);
  });

  it("extracts named ranges when the resolver returns a range", () => {
    const input = "=SUM(SalesData)";
    const tokenStart = input.indexOf("SalesData");
    const tokenEnd = tokenStart + "SalesData".length;

    const { references } = extractFormulaReferences(input, 0, 0, {
      resolveName: (name) =>
        name === "SalesData" ? { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 0 } : null
    });

    expect(references).toEqual([
      {
        text: "SalesData",
        range: { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 0 },
        index: 0,
        start: tokenStart,
        end: tokenEnd
      }
    ]);
  });

  it("ignores unresolved identifiers so we don't highlight every name-like token", () => {
    const input = "=UnknownName + A1";
    const { references } = extractFormulaReferences(input, 0, 0, { resolveName: () => null });
    expect(references.map((r) => r.text)).toEqual(["A1"]);
  });

  it("detects activeIndex for named ranges at the caret (including token end)", () => {
    const input = "=SUM(SalesData)";
    const tokenStart = input.indexOf("SalesData");
    const tokenEnd = tokenStart + "SalesData".length;

    const resolveName = (name: string) =>
      name === "SalesData" ? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } : null;

    // Caret inside token.
    expect(extractFormulaReferences(input, tokenStart + 1, tokenStart + 1, { resolveName }).activeIndex).toBe(0);
    // Caret at end of token should still count as inside.
    expect(extractFormulaReferences(input, tokenEnd, tokenEnd, { resolveName }).activeIndex).toBe(0);
  });

  it("does not treat function-call identifiers as named ranges", () => {
    const input = "=MyFunc(MyRange)";
    const { references } = extractFormulaReferences(input, 0, 0, {
      resolveName: (name) => (name === "MyFunc" || name === "MyRange" ? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } : null),
    });
    // Only the argument identifier should be considered (the function name is tokenized as `function`).
    expect(references.map((r) => r.text)).toEqual(["MyRange"]);
  });

  it("does not treat identifiers followed by whitespace and '(' as named ranges", () => {
    const input = "=MyFunc (MyRange)";
    const { references } = extractFormulaReferences(input, 0, 0, {
      resolveName: (name) => (name === "MyFunc" || name === "MyRange" ? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } : null),
    });
    expect(references.map((r) => r.text)).toEqual(["MyRange"]);
  });

  it("does not treat TRUE/FALSE identifiers as named ranges", () => {
    const input = "=TRUE+FALSE+MyRange";
    const { references } = extractFormulaReferences(input, 0, 0, {
      resolveName: (name) => ({ startRow: 0, startCol: 0, endRow: 0, endCol: 0, sheet: name }),
    });
    expect(references.map((r) => r.text)).toEqual(["MyRange"]);
  });

  it("preserves stable reference indices when mixing A1 and named ranges", () => {
    const input = "=A1+SalesData+B2";
    const { references } = extractFormulaReferences(input, 0, 0, {
      resolveName: (name) =>
        name === "SalesData" ? { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 0 } : null,
    });
    expect(references.map((r) => [r.text, r.index])).toEqual([
      ["A1", 0],
      ["SalesData", 1],
      ["B2", 2],
    ]);
  });

  it("detects activeIndex correctly when mixing A1 and named ranges", () => {
    const input = "=A1+SalesData+B2";
    const start = input.indexOf("SalesData");
    const end = start + "SalesData".length;
    const resolveName = (name: string) =>
      name === "SalesData" ? { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 0 } : null;

    expect(extractFormulaReferences(input, start + 1, start + 1, { resolveName }).activeIndex).toBe(1);
    // Caret at end should still count as inside.
    expect(extractFormulaReferences(input, end, end, { resolveName }).activeIndex).toBe(1);
  });
  it("extracts structured table references (data rows only)", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=SUM(Table1[Amount])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[Amount]",
      range: { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 3, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[Amount]".length
    });
  });

  it("extracts structured table references with #All (includes header row)", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#All],[Amount]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#All],[Amount]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount]]".length
    });
  });

  it("extracts multi-column structured table references when columns are contiguous", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:D4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 3,
          columns: ["Item", "Amount", "Tax", "Total"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#All],[Amount],[Tax]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#All],[Amount],[Tax]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 2 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount],[Tax]]".length
    });
  });

  it("does not resolve non-contiguous multi-column structured refs into a misleading rectangular range", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 3,
          columns: ["Item", "Amount", "Tax", "Total"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#All],[Amount],[Total]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toEqual([]);
  });

  it("extracts structured table references with multi-column ranges (:) inside nested selectors", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:D4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 3,
          columns: ["Item", "Amount", "Tax", "Total"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#All],[Amount]:[Total]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#All],[Amount]:[Total]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 3 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount]:[Total]]".length
    });

    const dataInput = "=SUM(Table1[[Amount]:[Total]])";
    const { references: dataRefs } = extractFormulaReferences(dataInput, 0, 0, { tables });
    expect(dataRefs).toHaveLength(1);
    expect(dataRefs[0]).toEqual({
      text: "Table1[[Amount]:[Total]]",
      range: { sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 3, endCol: 3 },
      index: 0,
      start: dataInput.indexOf("Table1"),
      end: dataInput.indexOf("Table1") + "Table1[[Amount]:[Total]]".length
    });
  });

  it("extracts structured references even when they are followed by operators", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=Table1[[#All],[Amount]]+1";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#All],[Amount]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount]]".length
    });
  });

  it("extracts structured references even when followed by string literals containing `]]`", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = '=SUM(Table1[[#All],[Amount]] & "]]", 1)';
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#All],[Amount]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#All],[Amount]]".length
    });
  });

  it("extracts structured table references with explicit selectors (#Headers/#Data)", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const headersInput = "=SUM(Table1[[#Headers],[Amount]])";
    const { references: headersRefs } = extractFormulaReferences(headersInput, 0, 0, { tables });
    expect(headersRefs).toHaveLength(1);
    expect(headersRefs[0]).toEqual({
      text: "Table1[[#Headers],[Amount]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 0, endCol: 1 },
      index: 0,
      start: headersInput.indexOf("Table1"),
      end: headersInput.indexOf("Table1") + "Table1[[#Headers],[Amount]]".length
    });

    const dataInput = "=SUM(Table1[[#Data],[Amount]])";
    const { references: dataRefs } = extractFormulaReferences(dataInput, 0, 0, { tables });
    expect(dataRefs).toHaveLength(1);
    expect(dataRefs[0]?.text).toBe("Table1[[#Data],[Amount]]");
    expect(dataRefs[0]?.range).toEqual({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 3, endCol: 1 });

    const totalsInput = "=SUM(Table1[[#Totals],[Amount]])";
    const { references: totalsRefs } = extractFormulaReferences(totalsInput, 0, 0, { tables });
    expect(totalsRefs).toHaveLength(1);
    expect(totalsRefs[0]?.text).toBe("Table1[[#Totals],[Amount]]");
    expect(totalsRefs[0]?.range).toEqual({ sheet: "Sheet1", startRow: 3, startCol: 1, endRow: 3, endCol: 1 });
  });

  it("extracts multi-item structured refs when the selector union is a contiguous rectangle", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"],
        },
      ],
    ]);

    const input = "=SUM(Table1[[#Headers],[#Data],[Amount]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    // #Headers + #Data is equivalent to #All in rectangular form.
    expect(references[0]?.range).toEqual({ sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 1 });
  });

  it("extracts multi-item structured refs without columns when the selector union is rectangular", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"],
        },
      ],
    ]);

    const input = "=COUNTA(Table1[[#All],[#Totals]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    // #All already includes the totals row; the union should still resolve to the full table range.
    expect(references[0]?.range).toEqual({ sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 3, endCol: 1 });
  });

  it("does not resolve discontiguous multi-item selector unions into a misleading rectangle", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"],
        },
      ],
    ]);

    const input = "=SUM(Table1[[#Headers],[#Totals],[Amount]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    // Header + totals is a discontiguous row union; avoid highlighting a bounding rectangle.
    expect(references).toEqual([]);
  });

  it("extracts contiguous multi-column unions that combine ranges and single columns", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:D4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 3,
          columns: ["Item", "Amount", "Tax", "Total"],
        },
      ],
    ]);

    const input = "=SUM(Table1[[#All],[Amount]:[Total],[Tax]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    // Amount:Total already includes Tax; the union is still a contiguous rectangle (B1:D4).
    expect(references[0]?.range).toEqual({ sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 3, endCol: 3 });
  });

  it("extracts structured table references with escaped closing brackets in column names", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "A]B"]
        }
      ]
    ]);

    // Excel escapes `]` inside structured reference items by doubling it: `]]`.
    const input = "=COUNTA(Table1[[#Headers],[A]]B]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#Headers],[A]]B]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 0, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#Headers],[A]]B]]".length
    });
  });

  it("extracts structured refs where escaped `]` is followed by operator characters inside the column name", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "A]+B"]
        }
      ]
    ]);

    const input = "=COUNTA(Table1[[#Headers],[A]]+B]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]).toEqual({
      text: "Table1[[#Headers],[A]]+B]]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 1, endRow: 0, endCol: 1 },
      index: 0,
      start: input.indexOf("Table1"),
      end: input.indexOf("Table1") + "Table1[[#Headers],[A]]+B]]".length
    });
  });

  it("does not resolve structured refs with unsupported selectors (e.g. #This Row)", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#This Row],[Amount]])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toEqual([]);
  });

  it("extracts structured table specifiers like #All/#Headers/#Data", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Full table range (including header row) is A1:B4 in Excel terms.
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const allInput = "=SUM(Table1[#All])";
    const { references: allRefs } = extractFormulaReferences(allInput, 0, 0, { tables });
    expect(allRefs).toHaveLength(1);
    expect(allRefs[0]).toEqual({
      text: "Table1[#All]",
      range: { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
      index: 0,
      start: allInput.indexOf("Table1"),
      end: allInput.indexOf("Table1") + "Table1[#All]".length
    });

    const headersInput = "=SUM(Table1[#Headers])";
    const { references: headerRefs } = extractFormulaReferences(headersInput, 0, 0, { tables });
    expect(headerRefs).toHaveLength(1);
    expect(headerRefs[0]?.range).toEqual({ sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 });

    const totalsInput = "=SUM(Table1[#Totals])";
    const { references: totalsRefs } = extractFormulaReferences(totalsInput, 0, 0, { tables });
    expect(totalsRefs).toHaveLength(1);
    expect(totalsRefs[0]?.range).toEqual({ sheet: "Sheet1", startRow: 3, startCol: 0, endRow: 3, endCol: 1 });

    const dataInput = "=SUM(Table1[#Data])";
    const { references: dataRefs } = extractFormulaReferences(dataInput, 0, 0, { tables });
    expect(dataRefs).toHaveLength(1);
    expect(dataRefs[0]?.range).toEqual({ sheet: "Sheet1", startRow: 1, startCol: 0, endRow: 3, endCol: 1 });
  });

  it("treats a structured ref as the active reference even when the caret is after an internal comma", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=SUM(Table1[[#All],[Amount]])";
    const comma = input.indexOf(",");
    expect(comma).toBeGreaterThan(0);

    const { activeIndex: inside } = extractFormulaReferences(input, comma + 1, comma + 1, { tables });
    expect(inside).toBe(0);

    const { activeIndex: outside } = extractFormulaReferences(input, input.length, input.length, { tables });
    expect(outside).toBe(null);
  });

  it("resolves structured refs even when table coordinates are reversed", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          // Same table as A1:B4, but with reversed start/end coordinates.
          startRow: 3,
          startCol: 1,
          endRow: 0,
          endCol: 0,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = "=SUM(Table1[Amount])";
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(1);
    expect(references[0]?.range).toEqual({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 3, endCol: 1 });
  });

  it("does not extract structured refs inside string literals", () => {
    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Amount"]
        }
      ]
    ]);

    const input = '=SUM("Table1[Amount]")';
    const { references } = extractFormulaReferences(input, 0, 0, { tables });
    expect(references).toHaveLength(0);
  });
});

describe("assignFormulaReferenceColors", () => {
  it("assigns palette colors by index on first pass", () => {
    const { references } = extractFormulaReferences("=A1+B1", 0, 0);
    const { colored } = assignFormulaReferenceColors(references, null);
    expect(colored.map((r) => r.color)).toEqual([FORMULA_REFERENCE_PALETTE[0], FORMULA_REFERENCE_PALETTE[1]]);
  });

  it("reuses the same color for repeated references within a formula", () => {
    const { references } = extractFormulaReferences("=A1+A1", 0, 0);
    const { colored } = assignFormulaReferenceColors(references, null);
    expect(colored).toHaveLength(2);
    expect(colored[0]?.color).toBe(FORMULA_REFERENCE_PALETTE[0]);
    expect(colored[1]?.color).toBe(FORMULA_REFERENCE_PALETTE[0]);
  });

  it("reuses colors for the same reference text across edits", () => {
    const first = extractFormulaReferences("=A1+B1", 0, 0).references;
    const { colored: coloredFirst, nextByText } = assignFormulaReferenceColors(first, null);

    const second = extractFormulaReferences("=B1+A1", 0, 0).references;
    const { colored: coloredSecond } = assignFormulaReferenceColors(second, nextByText);

    expect(coloredFirst.map((r) => [r.text, r.color])).toEqual([
      ["A1", FORMULA_REFERENCE_PALETTE[0]],
      ["B1", FORMULA_REFERENCE_PALETTE[1]]
    ]);
    expect(coloredSecond.map((r) => [r.text, r.color])).toEqual([
      ["B1", FORMULA_REFERENCE_PALETTE[1]],
      ["A1", FORMULA_REFERENCE_PALETTE[0]]
    ]);
  });

  it("preserves existing reference colors when a new reference is inserted earlier", () => {
    const initialRefs = extractFormulaReferences("=A1+B1", 0, 0).references;
    const { nextByText } = assignFormulaReferenceColors(initialRefs, null);

    const editedRefs = extractFormulaReferences("=C1+A1+B1", 0, 0).references;
    const { colored } = assignFormulaReferenceColors(editedRefs, nextByText);

    expect(colored.map((r) => [r.text, r.color])).toEqual([
      ["C1", FORMULA_REFERENCE_PALETTE[2]],
      ["A1", FORMULA_REFERENCE_PALETTE[0]],
      ["B1", FORMULA_REFERENCE_PALETTE[1]]
    ]);
  });
});
