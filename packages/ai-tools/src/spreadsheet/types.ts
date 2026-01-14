export type CellScalar = string | number | boolean | null;

/**
 * Spreadsheet cell values as returned by host workbook backends.
 *
 * Tool schemas (e.g. `read_range`) only expose {@link CellScalar} values to the LLM, but
 * real spreadsheet engines can surface richer runtime values (entities/records, rich text,
 * in-cell images, etc.). `ToolExecutor` normalizes these rich values into safe scalar
 * representations (e.g. trimming in-cell image alt text).
 */
export type CellValue = unknown;

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  font_size?: number;
  font_color?: string;
  background_color?: string;
  number_format?: string;
  horizontal_align?: "left" | "center" | "right";
}

export interface CellData {
  value: CellValue;
  /**
   * Formula string including leading "=".
   *
   * Note:
   * - The in-memory workbook does not evaluate formulas, so formula cells typically have `value: null`.
   * - Real spreadsheet backends may provide a computed `value` *alongside* `formula`. ToolExecutor
   *   can optionally surface/use those computed values when `include_formula_values` is enabled
   *   (with conservative DLP gating).
   */
  formula?: string;
  format?: CellFormat;
}

export function isCellEmpty(cell: CellData): boolean {
  return (cell.value === null || cell.value === undefined) && !cell.formula && (!cell.format || Object.keys(cell.format).length === 0);
}

export function cloneCell(cell: CellData): CellData {
  return {
    value: cell.value ?? null,
    formula: cell.formula,
    format: cell.format ? { ...cell.format } : undefined
  };
}
