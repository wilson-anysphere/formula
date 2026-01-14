import { describe, expect, it } from "vitest";

import { formatValueWithNumberFormat, isValidExcelNumberFormatCode } from "../numberFormat.js";

describe("formatValueWithNumberFormat", () => {
  it("formats time-only hh:mm:ss formats (Excel-style)", () => {
    expect(formatValueWithNumberFormat(0, "hh:mm:ss")).toBe("00:00:00");
    expect(formatValueWithNumberFormat(0.5, "hh:mm:ss")).toBe("12:00:00");
  });

  it("formats basic scientific notation presets", () => {
    expect(formatValueWithNumberFormat(1234.5, "0.00E+00")).toBe("1.23E+03");
    expect(formatValueWithNumberFormat(0.01234, "0.00E+00")).toBe("1.23E-02");
  });

  it("formats basic fraction presets", () => {
    expect(formatValueWithNumberFormat(1.5, "# ?/?")).toBe("1 1/2");
    expect(formatValueWithNumberFormat(0.3333333, "# ?/?")).toBe("1/3");
  });
});

describe("isValidExcelNumberFormatCode", () => {
  it("accepts common Excel-style custom number format codes", () => {
    expect(isValidExcelNumberFormatCode("0.00")).toBe(true);
    expect(isValidExcelNumberFormatCode("#,##0")).toBe(true);
    expect(isValidExcelNumberFormatCode("$#,##0.00")).toBe(true);
    expect(isValidExcelNumberFormatCode('";"0.00')).toBe(true);
    expect(isValidExcelNumberFormatCode('0.00;[Red]-0.00')).toBe(true);
  });

  it("rejects obvious syntax errors (unbalanced quotes/brackets, dangling escapes, too many sections)", () => {
    expect(isValidExcelNumberFormatCode('"0.00')).toBe(false);
    expect(isValidExcelNumberFormatCode("[Red")).toBe(false);
    expect(isValidExcelNumberFormatCode("]0.00")).toBe(false);
    expect(isValidExcelNumberFormatCode("0;0;0;0;0")).toBe(false);
    expect(isValidExcelNumberFormatCode("0\\")).toBe(false);
    expect(isValidExcelNumberFormatCode("0*")).toBe(false);
    expect(isValidExcelNumberFormatCode("0_")).toBe(false);
  });
});
