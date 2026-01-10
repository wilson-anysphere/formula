export type SortOrder = "ascending" | "descending";

export type SortKey = {
  column: number;
  order: SortOrder;
};

export type SortSpec = {
  keys: SortKey[];
  hasHeader: boolean;
};

export type TextMatchKind = "contains" | "beginsWith" | "endsWith";

export type FilterCriterion =
  | { type: "equals"; value: string }
  | { type: "textMatch"; kind: TextMatchKind; pattern: string }
  | { type: "numberGreaterThan"; value: number }
  | { type: "numberLessThan"; value: number }
  | { type: "between"; min: number; max: number }
  | { type: "blanks" }
  | { type: "nonBlanks" };

export type ColumnFilter = {
  criteria: FilterCriterion[];
  join: "any" | "all";
};

export type AutoFilter = {
  rangeA1: string;
  columns: Record<number, ColumnFilter>;
};

export type FilterViewId = string;

