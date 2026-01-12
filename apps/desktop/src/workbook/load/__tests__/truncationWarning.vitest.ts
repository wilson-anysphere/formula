import { describe, expect, it, vi } from "vitest";

import { createWorkbookLoadTruncationWarning, warnIfWorkbookLoadTruncated } from "../truncationWarning.js";

describe("createWorkbookLoadTruncationWarning", () => {
  it("returns null when there are no truncations", () => {
    expect(createWorkbookLoadTruncationWarning([], { maxRows: 10_000, maxCols: 200 })).toBeNull();
  });

  it("includes sheet details + adjustment hint", () => {
    const message = createWorkbookLoadTruncationWarning(
      [
        {
          sheetId: "Sheet1",
          sheetName: "Foo",
          originalRange: { start_row: 0, end_row: 15_000, start_col: 0, end_col: 250 },
          loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 199 },
          truncatedRows: true,
          truncatedCols: true,
        },
      ],
      { maxRows: 10_000, maxCols: 200 },
    );

    expect(message).toContain("Workbook partially loaded");
    expect(message).toContain("Foo");
    expect(message).toContain("rows 1-10,000");
    expect(message).toContain("cols 1-200");
    expect(message).toContain("rows 1-15,001");
    expect(message).toContain("cols 1-251");
    expect(message).toContain("loadMaxRows");
    expect(message).toContain("DESKTOP_LOAD_MAX_ROWS");
  });
});

describe("warnIfWorkbookLoadTruncated", () => {
  it("emits a toast when truncation occurs", () => {
    const consoleWarn = vi.spyOn(console, "warn").mockImplementation(() => {});
    const showToast = vi.fn();
    try {
      warnIfWorkbookLoadTruncated(
        [
          {
            sheetId: "Sheet1",
            sheetName: "Foo",
            originalRange: { start_row: 0, end_row: 15_000, start_col: 0, end_col: 250 },
            loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 199 },
            truncatedRows: true,
            truncatedCols: true,
          },
        ],
        { maxRows: 10_000, maxCols: 200 },
        showToast,
      );
    } finally {
      consoleWarn.mockRestore();
    }

    expect(showToast).toHaveBeenCalledTimes(1);
    expect(showToast.mock.calls[0]?.[1]).toBe("warning");
    expect(showToast.mock.calls[0]?.[0]).toContain("Foo");
  });
});
