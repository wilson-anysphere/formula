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
    expect(message).toContain("VITE_DESKTOP_LOAD_MAX_ROWS");
  });

  it("summarizes multiple sheets and collapses overflow into +N more", () => {
    const message = createWorkbookLoadTruncationWarning(
      [
        {
          sheetId: "s1",
          sheetName: "Sheet 1",
          originalRange: { start_row: 0, end_row: 20_000, start_col: 0, end_col: 250 },
          loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 199 },
          truncatedRows: true,
          truncatedCols: true,
        },
        {
          sheetId: "s2",
          sheetName: "Sheet 2",
          originalRange: { start_row: 0, end_row: 11_000, start_col: 0, end_col: 10 },
          loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 10 },
          truncatedRows: true,
          truncatedCols: false,
        },
        {
          sheetId: "s3",
          sheetName: "Sheet 3",
          originalRange: { start_row: 0, end_row: 9, start_col: 0, end_col: 400 },
          loadedRange: { startRow: 0, endRow: 9, startCol: 0, endCol: 199 },
          truncatedRows: false,
          truncatedCols: true,
        },
        {
          sheetId: "s4",
          sheetName: "Sheet 4",
          originalRange: { start_row: 0, end_row: 50_000, start_col: 0, end_col: 500 },
          loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 199 },
          truncatedRows: true,
          truncatedCols: true,
        },
      ],
      { maxRows: 10_000, maxCols: 200 },
    );

    expect(message).toContain("Workbook partially loaded");
    expect(message).toContain("Sheet 1");
    expect(message).toContain("Sheet 2");
    expect(message).toContain("Sheet 3");
    expect(message).toContain("+1 more");
  });

  it("does not render NaN in the warning when usedRange values are invalid", () => {
    const message = createWorkbookLoadTruncationWarning(
      [
        {
          sheetId: "Sheet1",
          sheetName: "Foo",
          // Defensive: treat invalid backend ranges as 0-based (1-1 in user coords).
          originalRange: { start_row: Number.NaN as any, end_row: Number.NaN as any, start_col: Number.NaN as any, end_col: Number.NaN as any },
          loadedRange: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
          truncatedRows: true,
          truncatedCols: true,
        },
      ],
      { maxRows: 10_000, maxCols: 200 },
    );

    expect(message).not.toContain("NaN");
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

  it("emits a single toast for a workbook with multiple truncated sheets", () => {
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
          {
            sheetId: "Sheet2",
            sheetName: "Bar",
            originalRange: { start_row: 0, end_row: 25_000, start_col: 0, end_col: 10 },
            loadedRange: { startRow: 0, endRow: 9_999, startCol: 0, endCol: 10 },
            truncatedRows: true,
            truncatedCols: false,
          },
        ],
        { maxRows: 10_000, maxCols: 200 },
        showToast,
      );
    } finally {
      consoleWarn.mockRestore();
    }

    expect(showToast).toHaveBeenCalledTimes(1);
    expect(showToast.mock.calls[0]?.[0]).toContain("Foo");
    expect(showToast.mock.calls[0]?.[0]).toContain("Bar");
  });
});
