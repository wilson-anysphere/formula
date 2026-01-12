import type { SheetUsedRange } from "@formula/workbook-backend";

export const DEFAULT_DESKTOP_LOAD_MAX_ROWS = 10_000;
export const DEFAULT_DESKTOP_LOAD_MAX_COLS = 200;

export type WorkbookLoadLimits = Readonly<{
  maxRows: number;
  maxCols: number;
}>;

function parsePositiveInt(value: unknown): number | null {
  if (typeof value === "number") {
    if (!Number.isFinite(value) || !Number.isInteger(value) || value <= 0) return null;
    return value;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const parsed = Number(trimmed);
    if (!Number.isFinite(parsed) || !Number.isInteger(parsed) || parsed <= 0) return null;
    return parsed;
  }

  return null;
}

/**
 * Resolves the workbook load caps.
 *
 * Precedence:
 * 1) defaults
 * 2) env vars (DESKTOP_LOAD_MAX_ROWS/COLS)
 * 3) URL query params (?maxRows=…&maxCols=…)
 *
 * This function is intentionally pure so it can be unit-tested. Callers should
 * provide the query/env inputs from their environment (e.g. `window.location.search`,
 * `import.meta.env`, `process.env`).
 */
export function resolveWorkbookLoadLimits(
  options: Readonly<{
    queryString?: string | null;
    env?: Record<string, unknown> | null;
    defaults?: WorkbookLoadLimits;
  }> = {},
): WorkbookLoadLimits {
  const defaults = options.defaults ?? {
    maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS,
    maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS,
  };

  let maxRows = defaults.maxRows;
  let maxCols = defaults.maxCols;

  const env = options.env ?? {};
  const envMaxRows = parsePositiveInt(env.DESKTOP_LOAD_MAX_ROWS ?? env.VITE_DESKTOP_LOAD_MAX_ROWS);
  const envMaxCols = parsePositiveInt(env.DESKTOP_LOAD_MAX_COLS ?? env.VITE_DESKTOP_LOAD_MAX_COLS);
  if (envMaxRows != null) maxRows = envMaxRows;
  if (envMaxCols != null) maxCols = envMaxCols;

  const queryString = options.queryString ?? "";
  if (queryString) {
    const params = new URLSearchParams(queryString.startsWith("?") ? queryString.slice(1) : queryString);
    const queryMaxRows = parsePositiveInt(params.get("maxRows"));
    const queryMaxCols = parsePositiveInt(params.get("maxCols"));
    if (queryMaxRows != null) maxRows = queryMaxRows;
    if (queryMaxCols != null) maxCols = queryMaxCols;
  }

  return { maxRows, maxCols };
}

function clampNumber(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.max(min, Math.min(value, max));
}

export type ClampedUsedRangeResult = Readonly<{
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  truncatedRows: boolean;
  truncatedCols: boolean;
}>;

/**
 * Clamps a backend-provided usedRange to the configured caps, and reports whether
 * truncation happened.
 */
export function clampUsedRange(usedRange: SheetUsedRange, limits: WorkbookLoadLimits): ClampedUsedRangeResult {
  // Sanity-check caps, even though `resolveWorkbookLoadLimits` enforces positive ints.
  const maxRows = Number.isFinite(limits.maxRows) && limits.maxRows > 0 ? Math.floor(limits.maxRows) : DEFAULT_DESKTOP_LOAD_MAX_ROWS;
  const maxCols = Number.isFinite(limits.maxCols) && limits.maxCols > 0 ? Math.floor(limits.maxCols) : DEFAULT_DESKTOP_LOAD_MAX_COLS;

  const truncatedRows = usedRange.end_row >= maxRows;
  const truncatedCols = usedRange.end_col >= maxCols;

  const maxRowIndex = maxRows - 1;
  const maxColIndex = maxCols - 1;

  const startRow = clampNumber(usedRange.start_row, 0, maxRowIndex);
  const endRow = clampNumber(usedRange.end_row, 0, maxRowIndex);
  const startCol = clampNumber(usedRange.start_col, 0, maxColIndex);
  const endCol = clampNumber(usedRange.end_col, 0, maxColIndex);

  return { startRow, endRow, startCol, endCol, truncatedRows, truncatedCols };
}

