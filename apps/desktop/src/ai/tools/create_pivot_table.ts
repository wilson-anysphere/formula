/**
 * AI tool hook: create a pivot table from structured parameters.
 *
 * In the full application this will:
 * 1. Build/refresh a pivot cache from `sourceRange`
 * 2. Create a pivot definition object
 * 3. Compute the pivot and write the output into the worksheet starting at `destination`
 * 4. Keep the pivot definition so it can be refreshed on data changes and exported to XLSX.
 *
 * This repo slice only provides the type-level contract and a small validation
 * layer so the orchestrator can wire it up to the engine.
 */

import type { PivotTableConfig, ValueField } from "../../panels/pivot-builder/types";

export interface CreatePivotTableParams {
  source_range: string; // e.g. "Sheet1!A1:F1000"
  destination: string; // e.g. "Sheet2!A1"
  rows?: string[];
  columns?: string[];
  values: Array<{ field: string; aggregation: ValueField["aggregation"]; name?: string }>;
  filters?: Array<{ field: string; allowed?: Array<string | number | boolean | null> }>;
}

export interface CreatePivotTableResult {
  config: PivotTableConfig;
}

export function createPivotTable(params: CreatePivotTableParams): CreatePivotTableResult {
  const values = params.values.map((v) => ({
    sourceField: v.field,
    name: v.name ?? `${v.aggregation} of ${v.field}`,
    aggregation: v.aggregation,
  }));

  return {
    config: {
      rowFields: (params.rows ?? []).map((f) => ({ sourceField: f })),
      columnFields: (params.columns ?? []).map((f) => ({ sourceField: f })),
      valueFields: values,
      filterFields: (params.filters ?? []).map((f) => ({
        sourceField: f.field,
        allowed: f.allowed,
      })),
      layout: "tabular",
      subtotals: "none",
      grandTotals: { rows: true, columns: true },
    },
  };
}

