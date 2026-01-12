import { describe, expect, it } from "vitest";

import { formatSheetNameForA1, isValidUnquotedSheetNameForA1 } from "./formatSheetNameForA1";

describe("formatSheetNameForA1", () => {
  it("leaves simple sheet names unquoted", () => {
    expect(formatSheetNameForA1("Sheet1")).toBe("Sheet1");
    expect(formatSheetNameForA1("sheet_2")).toBe("sheet_2");
    expect(formatSheetNameForA1("Sheet.Name")).toBe("Sheet.Name");
  });

  it("quotes sheet names containing spaces/special characters", () => {
    expect(formatSheetNameForA1("My Sheet")).toBe("'My Sheet'");
    expect(formatSheetNameForA1("Sheet-1")).toBe("'Sheet-1'");
  });

  it("escapes apostrophes when quoting", () => {
    expect(formatSheetNameForA1("O'Brien")).toBe("'O''Brien'");
  });

  it("quotes reserved/ambiguous names (TRUE/FALSE)", () => {
    expect(formatSheetNameForA1("TRUE")).toBe("'TRUE'");
    expect(formatSheetNameForA1("false")).toBe("'false'");
  });

  it("quotes A1-style cell reference names (e.g. A1, XFD1048576)", () => {
    expect(formatSheetNameForA1("A1")).toBe("'A1'");
    expect(formatSheetNameForA1("XFD1048576")).toBe("'XFD1048576'");
    // Beyond the max Excel column (XFD = 16384) should not be considered a cell reference.
    expect(formatSheetNameForA1("XFE1")).toBe("XFE1");
  });

  it("quotes R1C1-style cell reference names (e.g. R, C, RC, R1C1)", () => {
    expect(formatSheetNameForA1("R")).toBe("'R'");
    expect(formatSheetNameForA1("C")).toBe("'C'");
    expect(formatSheetNameForA1("RC")).toBe("'RC'");
    expect(formatSheetNameForA1("R1C1")).toBe("'R1C1'");
  });

  it("quotes names starting with digits", () => {
    expect(formatSheetNameForA1("1Sheet")).toBe("'1Sheet'");
  });
});

describe("isValidUnquotedSheetNameForA1", () => {
  it("matches the expected conservative ASCII-only unquoted rules", () => {
    expect(isValidUnquotedSheetNameForA1("Sheet1")).toBe(true);
    expect(isValidUnquotedSheetNameForA1("Sheet.Name")).toBe(true);
    expect(isValidUnquotedSheetNameForA1("My Sheet")).toBe(false);
    expect(isValidUnquotedSheetNameForA1("TRUE")).toBe(false);
    expect(isValidUnquotedSheetNameForA1("A1")).toBe(false);
    expect(isValidUnquotedSheetNameForA1("R1C1")).toBe(false);
  });
});

