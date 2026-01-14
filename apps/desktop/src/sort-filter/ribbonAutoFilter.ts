import { parseA1Range } from "../../../../packages/search/index.js";

export type RibbonFilterRange = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export type RibbonFilterColumn = {
  /**
   * 0-based column index *within the filter range* (i.e. relative to `range.startCol`).
   */
  colId: number;
  /**
   * Allowed values for this column. Values are compared via exact string match.
   */
  values: string[];
};

export type RibbonAutoFilterState = {
  /**
   * The filter range (A1, unqualified; e.g. "A1:D10").
   */
  rangeA1: string;
  /**
   * Number of header rows at the top of `rangeA1` that should never be filtered.
   *
   * For the ribbon MVP we always use `1`, matching Excel's default AutoFilter behavior.
   */
  headerRows: number;
  filterColumns: RibbonFilterColumn[];
};

type StoredRibbonAutoFilter = RibbonAutoFilterState & {
  range: RibbonFilterRange;
};

function clampNonNegativeInt(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.trunc(value));
}

function normalizeRange(range: RibbonFilterRange): RibbonFilterRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function parseRangeA1(rangeA1: string): RibbonFilterRange {
  return normalizeRange(parseA1Range(rangeA1));
}

function normalizeColumnId(colId: number): number | null {
  if (!Number.isFinite(colId)) return null;
  const id = Math.trunc(colId);
  if (id < 0) return null;
  return id;
}

function normalizeFilterValues(values: string[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const v of values) {
    const s = String(v ?? "");
    if (seen.has(s)) continue;
    seen.add(s);
    out.push(s);
  }
  return out;
}

function normalizeFilterColumns(columns: RibbonFilterColumn[]): RibbonFilterColumn[] {
  const byId = new Map<number, string[]>();
  for (const col of columns) {
    const id = normalizeColumnId(col.colId);
    if (id == null) continue;
    byId.set(id, normalizeFilterValues(Array.isArray(col.values) ? col.values : []));
  }
  return Array.from(byId.entries())
    .sort(([a], [b]) => a - b)
    .map(([colId, values]) => ({ colId, values }));
}

export class RibbonAutoFilterStore {
  private readonly filtersBySheet = new Map<string, Map<string, StoredRibbonAutoFilter>>();

  clearAll(): void {
    this.filtersBySheet.clear();
  }

  hasAny(sheetId: string): boolean {
    const view = this.filtersBySheet.get(sheetId);
    return Boolean(view && view.size > 0);
  }

  list(sheetId: string): RibbonAutoFilterState[] {
    const view = this.filtersBySheet.get(sheetId);
    if (!view) return [];
    return Array.from(view.values()).map((f) => ({
      rangeA1: f.rangeA1,
      headerRows: f.headerRows,
      filterColumns: f.filterColumns,
    }));
  }

  get(sheetId: string, rangeA1: string): RibbonAutoFilterState | undefined {
    const view = this.filtersBySheet.get(sheetId);
    const f = view?.get(rangeA1);
    if (!f) return undefined;
    return { rangeA1: f.rangeA1, headerRows: f.headerRows, filterColumns: f.filterColumns };
  }

  /**
   * Find an AutoFilter whose range contains `cell` (inclusive).
   */
  findByCell(sheetId: string, cell: { row: number; col: number }): RibbonAutoFilterState | undefined {
    const view = this.filtersBySheet.get(sheetId);
    if (!view) return undefined;
    const row = Math.trunc(cell.row);
    const col = Math.trunc(cell.col);
    if (!Number.isFinite(row) || !Number.isFinite(col)) return undefined;

    for (const f of view.values()) {
      const r = f.range;
      if (row < r.startRow || row > r.endRow) continue;
      if (col < r.startCol || col > r.endCol) continue;
      return { rangeA1: f.rangeA1, headerRows: f.headerRows, filterColumns: f.filterColumns };
    }
    return undefined;
  }

  /**
   * Create or update a filter. Callers are responsible for applying the filter
   * to the sheet (e.g. setting outline hidden rows).
   */
  set(sheetId: string, filter: RibbonAutoFilterState): RibbonAutoFilterState {
    const rangeA1 = String(filter.rangeA1 ?? "").trim();
    if (!rangeA1) {
      throw new Error("RibbonAutoFilterStore.set: rangeA1 is required");
    }

    const headerRows = clampNonNegativeInt(filter.headerRows);
    const stored: StoredRibbonAutoFilter = {
      rangeA1,
      headerRows,
      filterColumns: normalizeFilterColumns(filter.filterColumns ?? []),
      range: parseRangeA1(rangeA1),
    };

    let view = this.filtersBySheet.get(sheetId);
    if (!view) {
      view = new Map();
      this.filtersBySheet.set(sheetId, view);
    }
    view.set(rangeA1, stored);
    return { rangeA1: stored.rangeA1, headerRows: stored.headerRows, filterColumns: stored.filterColumns };
  }

  /**
   * Upsert a single column's filter values within an existing range.
   */
  setColumn(sheetId: string, rangeA1: string, args: { headerRows: number; colId: number; values: string[] }): RibbonAutoFilterState {
    const existing = this.filtersBySheet.get(sheetId)?.get(rangeA1);
    const colId = normalizeColumnId(args.colId);
    if (colId == null) {
      throw new Error("RibbonAutoFilterStore.setColumn: invalid colId");
    }

    const nextColumns = (() => {
      const current = existing?.filterColumns ?? [];
      const normalizedValues = normalizeFilterValues(args.values ?? []);
      const updated = current.filter((c) => c.colId !== colId);
      updated.push({ colId, values: normalizedValues });
      return normalizeFilterColumns(updated);
    })();

    return this.set(sheetId, { rangeA1, headerRows: args.headerRows, filterColumns: nextColumns });
  }

  delete(sheetId: string, rangeA1: string): void {
    const view = this.filtersBySheet.get(sheetId);
    if (!view) return;
    view.delete(rangeA1);
    if (view.size === 0) this.filtersBySheet.delete(sheetId);
  }

  clearSheet(sheetId: string): void {
    this.filtersBySheet.delete(sheetId);
  }
}

export function computeUniqueFilterValues(args: {
  range: RibbonFilterRange;
  headerRows: number;
  /**
   * Column index *within the range* (0-based relative to `range.startCol`).
   */
  colId: number;
  getValue: (row: number, col: number) => string;
}): string[] {
  const range = normalizeRange(args.range);
  const headerRows = clampNonNegativeInt(args.headerRows);
  const colId = normalizeColumnId(args.colId);
  if (colId == null) return [];

  const dataStartRow = range.startRow + headerRows;
  if (dataStartRow > range.endRow) return [];

  const absCol = range.startCol + colId;
  if (absCol < range.startCol || absCol > range.endCol) return [];

  const set = new Set<string>();
  for (let row = dataStartRow; row <= range.endRow; row += 1) {
    set.add(args.getValue(row, absCol));
  }
  return Array.from(set).sort((a, b) => {
    // Match Excel-like ordering where blanks appear last.
    if (a === "" && b === "") return 0;
    if (a === "") return 1;
    if (b === "") return -1;
    return a.localeCompare(b);
  });
}

export function computeFilterHiddenRows(args: {
  range: RibbonFilterRange;
  headerRows: number;
  filterColumns: RibbonFilterColumn[];
  getValue: (row: number, col: number) => string;
}): number[] {
  const range = normalizeRange(args.range);
  const headerRows = clampNonNegativeInt(args.headerRows);
  const columns = normalizeFilterColumns(args.filterColumns ?? []);
  if (columns.length === 0) return [];

  const dataStartRow = range.startRow + headerRows;
  if (dataStartRow > range.endRow) return [];

  const normalizedColumns = columns
    .map((c) => {
      const absCol = range.startCol + c.colId;
      if (absCol < range.startCol || absCol > range.endCol) return null;
      return { absCol, allowed: new Set(c.values.map((v) => String(v ?? ""))) };
    })
    .filter((c): c is NonNullable<typeof c> => c !== null);

  if (normalizedColumns.length === 0) return [];

  const hidden: number[] = [];
  for (let row = dataStartRow; row <= range.endRow; row += 1) {
    let ok = true;
    for (const col of normalizedColumns) {
      // Explicit empty selection: hide all data rows.
      if (col.allowed.size === 0) {
        ok = false;
        break;
      }
      const value = args.getValue(row, col.absCol);
      if (!col.allowed.has(value)) {
        ok = false;
        break;
      }
    }
    if (!ok) hidden.push(row);
  }
  return hidden;
}
