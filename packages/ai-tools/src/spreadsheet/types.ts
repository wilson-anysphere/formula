export type CellScalar = string | number | boolean | null;

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
  value: CellScalar;
  /**
   * Formula string including leading "=".
   *
   * Note: The in-memory model does not evaluate formulas; formula cells are stored
   * separately from `value` so tools can distinguish between literal strings and formulas.
   */
  formula?: string;
  format?: CellFormat;
}

export function isCellEmpty(cell: CellData): boolean {
  return cell.value === null && !cell.formula && (!cell.format || Object.keys(cell.format).length === 0);
}

export function cloneCell(cell: CellData): CellData {
  return {
    value: cell.value,
    formula: cell.formula,
    format: cell.format ? { ...cell.format } : undefined
  };
}

