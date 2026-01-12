import { describe, expect, it } from "vitest";

import { getOpenFileFilters } from "./file_dialog_filters.js";

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
});
