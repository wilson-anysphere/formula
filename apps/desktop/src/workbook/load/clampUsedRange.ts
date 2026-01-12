import type { SheetUsedRange } from "@formula/workbook-backend";

export const DEFAULT_DESKTOP_LOAD_MAX_ROWS = 10_000;
export const DEFAULT_DESKTOP_LOAD_MAX_COLS = 200;
export const DEFAULT_DESKTOP_LOAD_CHUNK_ROWS = 200;

export const WORKBOOK_LOAD_MAX_ROWS_STORAGE_KEY = "formula.desktop.workbookLoadMaxRows";
export const WORKBOOK_LOAD_MAX_COLS_STORAGE_KEY = "formula.desktop.workbookLoadMaxCols";
export const WORKBOOK_LOAD_CHUNK_ROWS_STORAGE_KEY = "formula.desktop.workbookLoadChunkRows";

export type WorkbookLoadLimits = Readonly<{
  maxRows: number;
  maxCols: number;
}>;

export type WorkbookLoadLimitOverrides = Readonly<{
  maxRows?: unknown;
  maxCols?: unknown;
  chunkRows?: unknown;
}>;

function parsePositiveInt(value: unknown): number | null {
  if (typeof value === "number") {
    if (!Number.isFinite(value) || !Number.isInteger(value) || value <= 0) return null;
    return value;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    // Allow common digit separators ("10_000", "10,000") since limits are often
    // copy/pasted from UI strings.
    const normalized = trimmed.replace(/[,_]/g, "");
    const parsed = Number(normalized);
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
 * 2) env vars (DESKTOP_LOAD_MAX_ROWS/COLS; VITE_DESKTOP_LOAD_MAX_ROWS/COLS)
 * 3) local overrides (e.g. values read from localStorage)
 * 4) URL query params (?loadMaxRows=…&loadMaxCols=…)
 *
 * Note: for backwards-compatibility we also accept `?maxRows=…&maxCols=…`.
 *
 * This function is intentionally pure so it can be unit-tested. Callers should
 * provide the query/env inputs from their environment (e.g. `window.location.search`,
 * `import.meta.env`, `process.env`).
 */
export function resolveWorkbookLoadLimits(
  options: Readonly<{
    queryString?: string | null;
    env?: Record<string, unknown> | null;
    overrides?: WorkbookLoadLimitOverrides | null;
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
  // Prefer the `DESKTOP_LOAD_*` variables, but fall back to `VITE_DESKTOP_LOAD_*` if the
  // preferred keys are missing or invalid.
  for (const candidate of [env.DESKTOP_LOAD_MAX_ROWS, env.VITE_DESKTOP_LOAD_MAX_ROWS]) {
    const parsed = parsePositiveInt(candidate);
    if (parsed != null) {
      maxRows = parsed;
      break;
    }
  }
  for (const candidate of [env.DESKTOP_LOAD_MAX_COLS, env.VITE_DESKTOP_LOAD_MAX_COLS]) {
    const parsed = parsePositiveInt(candidate);
    if (parsed != null) {
      maxCols = parsed;
      break;
    }
  }

  const overrides = options.overrides ?? {};
  const overrideRows = parsePositiveInt(overrides.maxRows);
  const overrideCols = parsePositiveInt(overrides.maxCols);
  if (overrideRows != null) maxRows = overrideRows;
  if (overrideCols != null) maxCols = overrideCols;

  const queryString = options.queryString ?? "";
  if (queryString) {
    const params = new URLSearchParams(queryString.startsWith("?") ? queryString.slice(1) : queryString);
    const queryMaxRows = parsePositiveInt(params.get("loadMaxRows") ?? params.get("maxRows"));
    const queryMaxCols = parsePositiveInt(params.get("loadMaxCols") ?? params.get("maxCols"));
    if (queryMaxRows != null) maxRows = queryMaxRows;
    if (queryMaxCols != null) maxCols = queryMaxCols;
  }

  return { maxRows, maxCols };
}

/**
 * Resolves the workbook snapshot load chunk size (rows per backend `getRange`).
 *
 * Precedence matches `resolveWorkbookLoadLimits`:
 * 1) defaults
 * 2) env vars (DESKTOP_LOAD_CHUNK_ROWS; VITE_DESKTOP_LOAD_CHUNK_ROWS)
 * 3) local overrides (e.g. values read from localStorage)
 * 4) URL query params (?loadChunkRows=…)
 *
 * Note: for backwards-compatibility we also accept `?chunkRows=…`.
 */
export function resolveWorkbookLoadChunkRows(
  options: Readonly<{
    queryString?: string | null;
    env?: Record<string, unknown> | null;
    override?: unknown;
    defaultChunkRows?: number;
  }> = {},
): number {
  const defaultChunkRows = options.defaultChunkRows ?? DEFAULT_DESKTOP_LOAD_CHUNK_ROWS;
  let chunkRows = defaultChunkRows;

  const env = options.env ?? {};
  for (const candidate of [env.DESKTOP_LOAD_CHUNK_ROWS, env.VITE_DESKTOP_LOAD_CHUNK_ROWS]) {
    const parsed = parsePositiveInt(candidate);
    if (parsed != null) {
      chunkRows = parsed;
      break;
    }
  }

  const override = parsePositiveInt(options.override);
  if (override != null) chunkRows = override;

  const queryString = options.queryString ?? "";
  if (queryString) {
    const params = new URLSearchParams(queryString.startsWith("?") ? queryString.slice(1) : queryString);
    const queryChunkRows = parsePositiveInt(params.get("loadChunkRows") ?? params.get("chunkRows"));
    if (queryChunkRows != null) chunkRows = queryChunkRows;
  }

  return chunkRows;
}

function normalizeIndex(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.floor(value);
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

  // Clamp using an intersection against the [0, maxIndex] window. This preserves empty
  // intersections (e.g. when `usedRange.start_row` is already beyond `maxRows`), which
  // allows callers to skip range fetches entirely.
  const startRow = Math.max(0, normalizeIndex(usedRange.start_row));
  const endRow = Math.min(normalizeIndex(usedRange.end_row), maxRowIndex);
  const startCol = Math.max(0, normalizeIndex(usedRange.start_col));
  const endCol = Math.min(normalizeIndex(usedRange.end_col), maxColIndex);

  return { startRow, endRow, startCol, endCol, truncatedRows, truncatedCols };
}
