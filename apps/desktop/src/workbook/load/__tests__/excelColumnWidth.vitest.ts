import { describe, expect, it } from "vitest";

import { excelColWidthCharsToPx, excelColWidthPxToChars } from "../excelColumnWidth";

describe("excel column width conversion", () => {
  it("converts Excel's default width (8.43 chars) to 64px deterministically", () => {
    expect(excelColWidthCharsToPx(8.43)).toBe(64);
  });

  it("converts whole-number character widths to pixels", () => {
    expect(excelColWidthCharsToPx(25)).toBe(180);
  });

  it("round-trips whole-number widths via px without drift", () => {
    const px = excelColWidthCharsToPx(25);
    expect(excelColWidthPxToChars(px)).toBeCloseTo(25, 12);
  });
});

