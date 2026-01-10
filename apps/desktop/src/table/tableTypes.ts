export interface TableStyleInfo {
  name: string;
  showFirstColumn: boolean;
  showLastColumn: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
}

export interface TableColumn {
  id: number;
  name: string;
}

export interface FilterColumn {
  colId: number; // 0-based
  values: string[]; // allowed values
}

export interface AutoFilter {
  range: string; // A1 range (e.g. "A1:D4")
  filterColumns: FilterColumn[];
}

export interface SortState {
  colId: number;
  descending: boolean;
}

export interface Table {
  id: number;
  name: string;
  displayName: string;
  range: string; // A1 range
  headerRowCount: number;
  totalsRowCount: number;
  columns: TableColumn[];
  style?: TableStyleInfo;
  autoFilter?: AutoFilter;
  sortState?: SortState;
}

export type CellValue = string | number | boolean | null;

