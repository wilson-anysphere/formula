import { describe, expect, it } from "vitest";

import { parseA1Range, rangeToA1 } from "./a1.js";

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
