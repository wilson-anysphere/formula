export type AggregationType =
  | "sum"
  | "count"
  | "average"
  | "min"
  | "max"
  | "product"
  | "countNumbers"
  | "stdDev"
  | "stdDevP"
  | "var"
  | "varP";

export interface PivotField {
  sourceField: string;
  name?: string;
}

export interface ValueField {
  sourceField: string;
  name: string;
  aggregation: AggregationType;
}

export interface FilterField {
  sourceField: string;
  allowed?: Array<string | number | boolean | null>;
}

export type PivotLayout = "compact" | "tabular";
export type PivotSubtotalPosition = "top" | "bottom" | "none";

export interface PivotGrandTotals {
  rows: boolean;
  columns: boolean;
}

export interface PivotTableConfig {
  rowFields: PivotField[];
  columnFields: PivotField[];
  valueFields: ValueField[];
  filterFields: FilterField[];
  layout: PivotLayout;
  subtotals: PivotSubtotalPosition;
  grandTotals: PivotGrandTotals;
}

export interface PivotTableDefinition extends PivotTableConfig {
  id: string;
  name: string;
  sourceRange: string; // e.g. "Sheet1!A1:C100"
  destination: string; // e.g. "Sheet2!A1"
}

