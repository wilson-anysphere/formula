import { describe, expect, it } from "vitest";

import { getOpenFileFilters, isOpenWorkbookPath } from "./file_dialog_filters.js";

describe("getOpenFileFilters", () => {
  it("includes all supported spreadsheet extensions", () => {
    const filters = getOpenFileFilters();
    const extensions = new Set(filters.flatMap((filter) => filter.extensions));

    const expected = [
      "xlsx",
      "xlsm",
      "xltx",
      "xltm",
      "xlam",
      "xls",
      "xlt",
      "xla",
      "xlsb",
      "csv",
      "parquet",
    ];
    for (const ext of expected) {
      expect(extensions.has(ext)).toBe(true);
    }
  });

  it("matches openable workbook paths by extension", () => {
    expect(isOpenWorkbookPath("/tmp/book.xlsx")).toBe(true);
    expect(isOpenWorkbookPath("C:\\Users\\me\\Book1.XLSM")).toBe(true);
    expect(isOpenWorkbookPath("/tmp/book.csv")).toBe(true);
    expect(isOpenWorkbookPath("/tmp/book.parquet")).toBe(true);
    expect(isOpenWorkbookPath("/tmp/book.png")).toBe(false);
    expect(isOpenWorkbookPath("/tmp/book")).toBe(false);
    expect(isOpenWorkbookPath("")).toBe(false);
  });
});
