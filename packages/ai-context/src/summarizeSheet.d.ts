import type { DataRegionSchema, SheetSchema, TableSchema } from "./schema.js";

export interface SummarizeSheetOptions {
  /**
   * Maximum number of tables to include in the summary.
   */
  maxTables?: number;
  /**
   * Maximum number of data regions to include in the summary.
   */
  maxRegions?: number;
  /**
   * Maximum number of headers to include per table.
   */
  maxHeadersPerTable?: number;
  /**
   * Maximum number of inferred column types to include per table.
   * Defaults to `maxHeadersPerTable`.
   */
  maxTypesPerTable?: number;
  /**
   * Maximum number of headers to include per region.
   */
  maxHeadersPerRegion?: number;
  /**
   * Maximum number of inferred column types to include per region.
   * Defaults to `maxHeadersPerRegion`.
   */
  maxTypesPerRegion?: number;
  /**
   * Include table summaries (default true).
   */
  includeTables?: boolean;
  /**
   * Include region summaries (default true).
   */
  includeRegions?: boolean;
  /**
   * Maximum number of named ranges to include in the summary.
   */
  maxNamedRanges?: number;
  /**
   * Include named range summaries (default true).
   */
  includeNamedRanges?: boolean;
}

/**
 * Produce a compact, deterministic, schema-first summary of a sheet.
 *
 * Intended for LLM context when sampling raw cell values is too expensive.
 */
export function summarizeSheetSchema(schema: SheetSchema, options?: SummarizeSheetOptions): string;

/**
 * Summarize either a table schema or a data region schema.
 */
export function summarizeRegion(
  schemaRegionOrTable: TableSchema | DataRegionSchema,
  options?: SummarizeSheetOptions,
): string;
