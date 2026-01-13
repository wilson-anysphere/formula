import type { WorkbookSchemaSummary } from "./workbookSchema.js";

export interface SummarizeWorkbookOptions {
  /** Maximum number of sheet names to include in the sheet-name list. */
  maxSheets?: number;
  /** Maximum number of tables to include. */
  maxTables?: number;
  /** Maximum number of named ranges to include. */
  maxNamedRanges?: number;
  /** Maximum number of headers to include per table. */
  maxHeadersPerTable?: number;
  /** Maximum number of types to include per table (defaults to `maxHeadersPerTable`). */
  maxTypesPerTable?: number;
  /** Include the sheet list line (default true). */
  includeSheets?: boolean;
  /** Include table lines (default true). */
  includeTables?: boolean;
  /** Include named range lines (default true). */
  includeNamedRanges?: boolean;
}

/**
 * Produce a compact, deterministic, schema-first summary of a workbook.
 */
export function summarizeWorkbookSchema(schema: WorkbookSchemaSummary, options?: SummarizeWorkbookOptions): string;
