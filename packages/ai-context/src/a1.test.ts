import { describe, expect, it } from "vitest";

import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, parseA1Range, rangeToA1 } from "./a1.js";

describe("A1 utilities (sheet quoting)", () => {
  it("roundtrips unquoted identifier-like sheet names", () => {
    const input = "Sheet1!A1:B2";
    const parsed = parseA1Range(input);
    expect(parsed).toEqual({
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
    });
    expect(rangeToA1(parsed)).toBe(input);
  });

  it("roundtrips Excel-style quoted sheet names", () => {
    const input = "'My Sheet'!A1:B2";
    const parsed = parseA1Range(input);
    expect(parsed.sheetName).toBe("My Sheet");
    expect(rangeToA1(parsed)).toBe(input);
  });

  it("roundtrips Excel-style quoted sheet names with embedded quotes", () => {
    const input = "'Bob''s Sheet'!A1";
    const parsed = parseA1Range(input);
    expect(parsed.sheetName).toBe("Bob's Sheet");
    expect(rangeToA1(parsed)).toBe(input);
  });

  it("accepts legacy unquoted sheet names with spaces", () => {
    const parsed = parseA1Range("My Sheet!A1");
    expect(parsed.sheetName).toBe("My Sheet");
    expect(rangeToA1(parsed)).toBe("'My Sheet'!A1");
  });

  it("rangeToA1 accepts already-quoted sheetName input (back-compat)", () => {
    expect(
      rangeToA1({
        sheetName: "'My Sheet'",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'My Sheet'!A1");

    expect(
      rangeToA1({
        sheetName: "'Bob''s Sheet'",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'Bob''s Sheet'!A1");
  });

  it("quotes reserved/ambiguous sheet names (TRUE, A1, R1C1, leading digits)", () => {
    expect(
      rangeToA1({
        sheetName: "TRUE",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'TRUE'!A1");

    expect(
      rangeToA1({
        sheetName: "A1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'A1'!A1");

    expect(
      rangeToA1({
        sheetName: "R1C1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'R1C1'!A1");

    expect(
      rangeToA1({
        sheetName: "1Sheet",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      }),
    ).toBe("'1Sheet'!A1");
  });
});

describe("A1 utilities (parsing variants)", () => {
  it("parses absolute refs with $ (cell + range) and canonicalizes output", () => {
    expect(parseA1Range("$A$1")).toEqual({
      sheetName: undefined,
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
    });
    expect(rangeToA1(parseA1Range("$A$1"))).toBe("A1");

    expect(rangeToA1(parseA1Range("Sheet1!$A$1:$B$2"))).toBe("Sheet1!A1:B2");
    expect(rangeToA1(parseA1Range("'My Sheet'!$a$1:$b$2"))).toBe("'My Sheet'!A1:B2");
  });

  it("accepts lower-case column letters", () => {
    expect(rangeToA1(parseA1Range("a1:b2"))).toBe("A1:B2");
    expect(rangeToA1(parseA1Range("sheet1!a1"))).toBe("sheet1!A1");
  });

  it("parses whole-column ranges (A:C) using Excel max row limits", () => {
    const parsed = parseA1Range("A:C");
    expect(parsed).toEqual({
      sheetName: undefined,
      startRow: 0,
      startCol: 0,
      endRow: EXCEL_MAX_ROWS - 1,
      endCol: 2,
    });
    expect(rangeToA1(parsed)).toBe("A:C");

    expect(rangeToA1(parseA1Range("Sheet1!c:a"))).toBe("Sheet1!A:C");
  });

  it("parses whole-row ranges (1:10) using Excel max column limits", () => {
    const parsed = parseA1Range("1:10");
    expect(parsed).toEqual({
      sheetName: undefined,
      startRow: 0,
      startCol: 0,
      endRow: 9,
      endCol: EXCEL_MAX_COLS - 1,
    });
    expect(rangeToA1(parsed)).toBe("1:10");

    expect(rangeToA1(parseA1Range("'My Sheet'!10:1"))).toBe("'My Sheet'!1:10");
  });
});

describe("A1 utilities (invalid inputs)", () => {
  it("rejects invalid A1 references", () => {
    expect(() => parseA1Range("")).toThrow();
    expect(() => parseA1Range("!A1")).toThrow();

    // Invalid cell refs.
    expect(() => parseA1Range("A0")).toThrow();
    expect(() => parseA1Range("XFE1")).toThrow(); // beyond XFD
    expect(() => parseA1Range("A1048577")).toThrow(); // beyond max row

    // Invalid row/col ranges.
    expect(() => parseA1Range("A")).toThrow();
    expect(() => parseA1Range("0:1")).toThrow();
    expect(() => parseA1Range("A:1")).toThrow();
    expect(() => parseA1Range("A1:B")).toThrow();
    expect(() => parseA1Range("A:")).toThrow();
    expect(() => parseA1Range(":A")).toThrow();
  });
});
