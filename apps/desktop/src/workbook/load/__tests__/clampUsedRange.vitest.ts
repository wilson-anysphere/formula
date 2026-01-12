import { describe, expect, it } from "vitest";

import {
  clampUsedRange,
  DEFAULT_DESKTOP_LOAD_CHUNK_ROWS,
  DEFAULT_DESKTOP_LOAD_MAX_COLS,
  DEFAULT_DESKTOP_LOAD_MAX_ROWS,
  resolveWorkbookLoadChunkRows,
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

  it("handles degenerate/empty used ranges (start > end) without truncation", () => {
    const result = clampUsedRange(
      { start_row: 10, end_row: 5, start_col: 3, end_col: 2 },
      { maxRows: 100, maxCols: 10 },
    );

    expect(result).toEqual({
      startRow: 10,
      endRow: 5,
      startCol: 3,
      endCol: 2,
      truncatedRows: false,
      truncatedCols: false,
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

  it("uses overrides when provided", () => {
    const limits = resolveWorkbookLoadLimits({
      env: { DESKTOP_LOAD_MAX_ROWS: "123", DESKTOP_LOAD_MAX_COLS: "456" },
      overrides: { maxRows: "111", maxCols: "222" },
    });

    expect(limits).toEqual({ maxRows: 111, maxCols: 222 });
  });

  it("prefers query params over overrides", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?loadMaxRows=500&loadMaxCols=600",
      overrides: { maxRows: "111", maxCols: "222" },
    });

    expect(limits).toEqual({ maxRows: 500, maxCols: 600 });
  });

  it("prefers query params over env vars", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?loadMaxRows=111&loadMaxCols=222",
      env: { DESKTOP_LOAD_MAX_ROWS: "123", DESKTOP_LOAD_MAX_COLS: "456" },
    });

    expect(limits).toEqual({ maxRows: 111, maxCols: 222 });
  });

  it("accepts digit separators in query params", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?loadMaxRows=10,000&loadMaxCols=1_000",
    });

    expect(limits).toEqual({ maxRows: 10_000, maxCols: 1_000 });
  });

  it("accepts whitespace separators in query params", () => {
    const limits = resolveWorkbookLoadLimits({
      queryString: "?loadMaxRows=10%20000&loadMaxCols=200",
    });

    expect(limits).toEqual({ maxRows: 10_000, maxCols: 200 });
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

describe("resolveWorkbookLoadChunkRows", () => {
  it("uses the default when not configured", () => {
    expect(resolveWorkbookLoadChunkRows()).toBe(DEFAULT_DESKTOP_LOAD_CHUNK_ROWS);
  });

  it("uses env vars when query params are not provided", () => {
    expect(resolveWorkbookLoadChunkRows({ env: { DESKTOP_LOAD_CHUNK_ROWS: "123" } })).toBe(123);
  });

  it("uses the override when provided", () => {
    expect(resolveWorkbookLoadChunkRows({ env: { DESKTOP_LOAD_CHUNK_ROWS: "123" }, override: "50" })).toBe(50);
  });

  it("prefers query params over overrides", () => {
    expect(resolveWorkbookLoadChunkRows({ queryString: "?loadChunkRows=10", override: "50" })).toBe(10);
  });

  it("accepts digit separators in query params", () => {
    expect(resolveWorkbookLoadChunkRows({ queryString: "?loadChunkRows=1,000" })).toBe(1_000);
  });

  it("accepts whitespace separators in query params", () => {
    expect(resolveWorkbookLoadChunkRows({ queryString: "?loadChunkRows=1%20000" })).toBe(1_000);
  });

  it("falls back to VITE_* env vars when DESKTOP_* value is invalid", () => {
    expect(
      resolveWorkbookLoadChunkRows({
        env: {
          DESKTOP_LOAD_CHUNK_ROWS: "nope",
          VITE_DESKTOP_LOAD_CHUNK_ROWS: "321",
        },
      }),
    ).toBe(321);
  });

  it("falls back to defaults when config values are invalid", () => {
    expect(
      resolveWorkbookLoadChunkRows({
        queryString: "?chunkRows=-5",
        env: { DESKTOP_LOAD_CHUNK_ROWS: "0" },
      }),
    ).toBe(DEFAULT_DESKTOP_LOAD_CHUNK_ROWS);
  });
});
