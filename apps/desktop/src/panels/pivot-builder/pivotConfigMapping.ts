import type { PivotTableConfig } from "./types";

/**
 * JSON shape expected by the Rust `formula_engine::pivot::PivotConfig` serde model.
 *
 * Note: The desktop Tauri command wrapper (`create_pivot_table`) nests this object
 * inside a request payload that uses snake_case keys, but the pivot config itself
 * uses `#[serde(rename_all = "camelCase")]`.
 */
export interface RustPivotField {
  sourceField: string;
  sortOrder?: "ascending" | "descending" | "manual";
  manualSort?: unknown[];
}

export interface RustValueField {
  sourceField: string;
  name: string;
  aggregation: PivotTableConfig["valueFields"][number]["aggregation"];
  showAs?: unknown;
  baseField?: string;
  baseItem?: string;
}

export interface RustFilterField {
  sourceField: string;
  // Filter values are represented in Rust as a `HashSet<PivotKeyPart>` which
  // is not currently representable losslessly in JS (it serializes numbers as
  // raw f64 bit-pattern u64s). For now the UI only sends filter field presence.
  allowed?: unknown;
}

export interface RustPivotConfig {
  rowFields: RustPivotField[];
  columnFields: RustPivotField[];
  valueFields: RustValueField[];
  filterFields: RustFilterField[];
  calculatedFields: unknown[];
  calculatedItems: unknown[];
  layout: PivotTableConfig["layout"];
  subtotals: PivotTableConfig["subtotals"];
  grandTotals: PivotTableConfig["grandTotals"];
}

export function toRustPivotConfig(config: PivotTableConfig): RustPivotConfig {
  return {
    rowFields: config.rowFields.map((f) => ({ sourceField: f.sourceField })),
    columnFields: config.columnFields.map((f) => ({ sourceField: f.sourceField })),
    valueFields: config.valueFields.map((f) => ({
      sourceField: f.sourceField,
      name: f.name,
      aggregation: f.aggregation,
    })),
    filterFields: config.filterFields.map((f) => ({
      sourceField: f.sourceField,
    })),
    calculatedFields: [],
    calculatedItems: [],
    layout: config.layout,
    subtotals: config.subtotals,
    grandTotals: config.grandTotals,
  };
}
