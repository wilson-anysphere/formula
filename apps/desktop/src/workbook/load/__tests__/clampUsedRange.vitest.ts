import { describe, expect, it } from "vitest";

import {
  clampUsedRange,
  DEFAULT_DESKTOP_LOAD_MAX_COLS,
  DEFAULT_DESKTOP_LOAD_MAX_ROWS,
  resolveWorkbookLoadLimits,
} from "../clampUsedRange.js";

describe("clampUsedRange", () => {
  it("does not truncate when used range is within limits", () => {
    const result = clampUsedRange(
      { start_row: 0, end_row: 99, start_col: 0, end_col: 9 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toEqual({
      startRow: 0,
      endRow: 99,
      startCol: 0,
      endCol: 9,
      truncatedRows: false,
      truncatedCols: false,
    });
  });

  it("detects and clamps row truncation", () => {
    const result = clampUsedRange(
      { start_row: 0, end_row: 100, start_col: 0, end_col: 9 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toMatchObject({
      startRow: 0,
      endRow: 99,
      startCol: 0,
      endCol: 9,
      truncatedRows: true,
      truncatedCols: false,
    });
  });

  it("detects and clamps col truncation", () => {
    const result = clampUsedRange(
      { start_row: 0, end_row: 9, start_col: 0, end_col: 10 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toMatchObject({
      startRow: 0,
      endRow: 9,
      startCol: 0,
      endCol: 9,
      truncatedRows: false,
      truncatedCols: true,
    });
  });

  it("detects and clamps both row + col truncation", () => {
    const result = clampUsedRange(
      { start_row: 0, end_row: 100, start_col: 0, end_col: 10 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toMatchObject({
      startRow: 0,
      endRow: 99,
      startCol: 0,
      endCol: 9,
      truncatedRows: true,
      truncatedCols: true,
    });
  });
});

describe("resolveWorkbookLoadLimits", () => {
  it("falls back to defaults when config values are invalid", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?maxRows=not-a-number&maxCols=-5",
      env: { DESKTOP_LOAD_MAX_ROWS: "nope", DESKTOP_LOAD_MAX_COLS: "0" },
    });

    expect(limits).toEqual({ maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS, maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS });
  });
});

