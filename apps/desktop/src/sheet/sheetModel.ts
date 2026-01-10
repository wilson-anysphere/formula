import type { CellCoord, Range } from "../selection/types";
import type { RichText } from "../grid/text/rich-text/types.js";

/**
 * Minimal in-memory sparse sheet model.
 *
 * This is *not* the long-term engine; it's just enough to validate selection,
 * navigation and editing semantics end-to-end.
 */
export class SheetModel {
  private cells = new Map<string, RichText>();

  getCellRichText(cell: CellCoord): RichText | null {
    return this.cells.get(key(cell)) ?? null;
  }

  getCellValue(cell: CellCoord): string {
    return this.cells.get(key(cell))?.text ?? "";
  }

  setCellValue(cell: CellCoord, value: string | RichText): void {
    const rich: RichText = typeof value === "string" ? { text: value, runs: [] } : value;
    if (rich.text === "") {
      this.cells.delete(key(cell));
      return;
    }
    this.cells.set(key(cell), rich);
  }

  isCellEmpty(cell: CellCoord): boolean {
    return (this.cells.get(key(cell))?.text ?? "") === "";
  }

  getUsedRange(): Range | null {
    if (this.cells.size === 0) return null;

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;

    for (const k of this.cells.keys()) {
      const [rowStr, colStr] = k.split(",");
      const row = Number(rowStr);
      const col = Number(colStr);
      if (Number.isNaN(row) || Number.isNaN(col)) continue;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }

    if (!Number.isFinite(minRow) || !Number.isFinite(minCol) || !Number.isFinite(maxRow) || !Number.isFinite(maxCol)) {
      return null;
    }

    return { startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol };
  }
}

function key(cell: CellCoord): string {
  return `${cell.row},${cell.col}`;
}
