import { CellValue, FilterColumn, SortState, Table } from "./tableTypes";

export interface TableViewRow {
  /** Original row index within the sheet (0-based). */
  row: number;
  values: CellValue[];
}

/**
 * Apply a table's AutoFilter + sort state to a set of table rows.
 *
 * This is intentionally "UI level" logic: it does not mutate sheet data. It
 * produces a stable row order for rendering.
 */
export function buildTableView(
  table: Table,
  rows: TableViewRow[],
): TableViewRow[] {
  let out = rows;
  if (table.autoFilter) {
    out = applyAutoFilter(table.autoFilter.filterColumns, out);
  }
  if (table.sortState) {
    out = applySort(table.sortState, out);
  }
  return out;
}

export function distinctColumnValues(
  rows: TableViewRow[],
  colId: number,
): string[] {
  const set = new Set<string>();
  for (const row of rows) {
    const v = row.values[colId];
    if (v == null) continue;
    set.add(String(v));
  }
  return Array.from(set).sort((a, b) => {
    // Match Excel-like ordering where blanks appear last.
    if (a === "" && b === "") return 0;
    if (a === "") return 1;
    if (b === "") return -1;
    return a.localeCompare(b);
  });
}

export function applyAutoFilter(
  filterColumns: FilterColumn[],
  rows: TableViewRow[],
): TableViewRow[] {
  if (filterColumns.length === 0) return rows;
  return rows.filter((row) => {
    for (const filter of filterColumns) {
      if (filter.values.length === 0) continue;
      const v = row.values[filter.colId];
      const vStr = v == null ? "" : String(v);
      if (!filter.values.includes(vStr)) return false;
    }
    return true;
  });
}

export function applySort(sort: SortState, rows: TableViewRow[]): TableViewRow[] {
  const { colId, descending } = sort;
  const dir = descending ? -1 : 1;
  return [...rows].sort((a, b) => {
    const av = a.values[colId];
    const bv = b.values[colId];
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    if (typeof av === "number" && typeof bv === "number") return dir * (av - bv);
    return dir * String(av).localeCompare(String(bv));
  });
}
