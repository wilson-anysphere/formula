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

  it("returns an empty intersection when used range starts beyond the cap", () => {
    const result = clampUsedRange(
      { start_row: 150, end_row: 160, start_col: 0, end_col: 9 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toMatchObject({
      // No rows in common with [0..99].
      startRow: 150,
      endRow: 99,
      startCol: 0,
      endCol: 9,
      truncatedRows: true,
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
  it("uses env vars when query params are not provided", () => {
    const limits = resolveWorkbookLoadLimits({
      env: { DESKTOP_LOAD_MAX_ROWS: "123", DESKTOP_LOAD_MAX_COLS: "456" },
    });

    expect(limits).toEqual({ maxRows: 123, maxCols: 456 });
  });

  it("prefers query params over env vars", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?maxRows=111&maxCols=222",
      env: { DESKTOP_LOAD_MAX_ROWS: "123", DESKTOP_LOAD_MAX_COLS: "456" },
    });

    expect(limits).toEqual({ maxRows: 111, maxCols: 222 });
  });

  it("falls back to VITE_* env vars when DESKTOP_* values are invalid", () => {
    const limits = resolveWorkbookLoadLimits({
      env: {
        DESKTOP_LOAD_MAX_ROWS: "not-a-number",
        DESKTOP_LOAD_MAX_COLS: "-1",
        VITE_DESKTOP_LOAD_MAX_ROWS: "321",
        VITE_DESKTOP_LOAD_MAX_COLS: "654",
      },
    });

    expect(limits).toEqual({ maxRows: 321, maxCols: 654 });
  });

  it("falls back per-field when some config values are invalid", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?maxRows=5000&maxCols=bad",
      env: { DESKTOP_LOAD_MAX_COLS: "300" },
    });

    // maxRows comes from query; maxCols falls back to env since query is invalid.
    expect(limits).toEqual({ maxRows: 5000, maxCols: 300 });
  });

  it("falls back to defaults when config values are invalid", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?maxRows=not-a-number&maxCols=-5",
      env: { DESKTOP_LOAD_MAX_ROWS: "nope", DESKTOP_LOAD_MAX_COLS: "0" },
    });

    expect(limits).toEqual({ maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS, maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS });
  });
});
