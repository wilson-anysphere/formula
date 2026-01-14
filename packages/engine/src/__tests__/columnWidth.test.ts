import { describe, expect, it } from "vitest";

import { excelColWidthCharsToPixels, pixelsToExcelColWidthChars } from "../columnWidth.ts";

describe("Excel column width conversion helpers", () => {
  it("converts Excel's default width (8.43 chars) to 64px", () => {
    expect(excelColWidthCharsToPixels(8.43)).toBe(64);
  });

  it("converts 64px to Excel's default width (8.43 chars)", () => {
    expect(pixelsToExcelColWidthChars(64)).toBe(8.43);
  });

  it("round-trips via pixels for common values", () => {
    const widths = [1, 2, 8.43, 10, 25];
    for (const width of widths) {
      const px = excelColWidthCharsToPixels(width);
      const roundTripped = pixelsToExcelColWidthChars(px);
      expect(excelColWidthCharsToPixels(roundTripped)).toBe(px);
    }
  });
});

