import type { CellRange } from "../model/CellProvider.ts";
import type { CanvasGridRenderer } from "../rendering/CanvasGridRenderer.ts";

export type FillMode = "copy" | "series" | "formulas";

export const DEFAULT_FILL_HANDLE_SIZE_PX = 8;

export type FillDragAxis = "vertical" | "horizontal";

export type FillDragCommit = {
  sourceRange: CellRange;
  targetRange: CellRange;
  mode: FillMode;
};

export type FillDragPreview = {
  axis: FillDragAxis;
  /**
   * Range of cells that will be written (excludes the source range).
   */
  targetRange: CellRange;
  /**
   * The resulting selection after applying the fill (source + target).
   */
  unionRange: CellRange;
};

export type RectLike = { x: number; y: number; width: number; height: number };

export function pointInRect(x: number, y: number, rect: RectLike): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

export function computeFillPreview(
  sourceRange: CellRange,
  pointerCell: { row: number; col: number },
  out?: FillDragPreview
): FillDragPreview | null {
  const srcTop = sourceRange.startRow;
  const srcBottomExclusive = sourceRange.endRow;
  const srcBottom = srcBottomExclusive - 1;
  const srcLeft = sourceRange.startCol;
  const srcRightExclusive = sourceRange.endCol;
  const srcRight = srcRightExclusive - 1;

  const row = pointerCell.row;
  const col = pointerCell.col;

  const rowExtension = row < srcTop ? row - srcTop : row > srcBottom ? row - srcBottom : 0;
  const colExtension = col < srcLeft ? col - srcLeft : col > srcRight ? col - srcRight : 0;

  if (rowExtension === 0 && colExtension === 0) return null;

  const axis: FillDragAxis =
    rowExtension !== 0 && colExtension !== 0 ? (Math.abs(rowExtension) >= Math.abs(colExtension) ? "vertical" : "horizontal") : rowExtension !== 0 ? "vertical" : "horizontal";

  if (axis === "vertical") {
    const unionStartRow = Math.min(srcTop, row);
    const unionEndRow = Math.max(srcBottomExclusive, row + 1);
    const unionStartCol = sourceRange.startCol;
    const unionEndCol = sourceRange.endCol;

    const targetStartRow = row >= srcBottomExclusive ? srcBottomExclusive : row < srcTop ? row : null;
    const targetEndRow = row >= srcBottomExclusive ? row + 1 : row < srcTop ? srcTop : null;

    if (targetStartRow == null || targetEndRow == null) return null;

    if (out) {
      out.axis = axis;
      out.unionRange.startRow = unionStartRow;
      out.unionRange.endRow = unionEndRow;
      out.unionRange.startCol = unionStartCol;
      out.unionRange.endCol = unionEndCol;
      out.targetRange.startRow = targetStartRow;
      out.targetRange.endRow = targetEndRow;
      out.targetRange.startCol = unionStartCol;
      out.targetRange.endCol = unionEndCol;
      return out;
    }

    const unionRange: CellRange = { startRow: unionStartRow, endRow: unionEndRow, startCol: unionStartCol, endCol: unionEndCol };
    const targetRange: CellRange = {
      startRow: targetStartRow,
      endRow: targetEndRow,
      startCol: unionStartCol,
      endCol: unionEndCol
    };
    return { axis, targetRange, unionRange };
  }

  const unionStartRow = sourceRange.startRow;
  const unionEndRow = sourceRange.endRow;
  const unionStartCol = Math.min(srcLeft, col);
  const unionEndCol = Math.max(srcRightExclusive, col + 1);

  const targetStartCol = col >= srcRightExclusive ? srcRightExclusive : col < srcLeft ? col : null;
  const targetEndCol = col >= srcRightExclusive ? col + 1 : col < srcLeft ? srcLeft : null;

  if (targetStartCol == null || targetEndCol == null) return null;

  if (out) {
    out.axis = axis;
    out.unionRange.startRow = unionStartRow;
    out.unionRange.endRow = unionEndRow;
    out.unionRange.startCol = unionStartCol;
    out.unionRange.endCol = unionEndCol;
    out.targetRange.startRow = unionStartRow;
    out.targetRange.endRow = unionEndRow;
    out.targetRange.startCol = targetStartCol;
    out.targetRange.endCol = targetEndCol;
    return out;
  }

  const unionRange: CellRange = { startRow: unionStartRow, endRow: unionEndRow, startCol: unionStartCol, endCol: unionEndCol };
  const targetRange: CellRange = { startRow: unionStartRow, endRow: unionEndRow, startCol: targetStartCol, endCol: targetEndCol };
  return { axis, targetRange, unionRange };
}

export function getSelectionHandleRect(renderer: CanvasGridRenderer, options?: { size?: number }): RectLike | null {
  void options;
  return renderer.getFillHandleRect();
}

export function hitTestSelectionHandle(renderer: CanvasGridRenderer, viewportX: number, viewportY: number): boolean {
  const rect = getSelectionHandleRect(renderer);
  if (!rect) return false;
  return pointInRect(viewportX, viewportY, rect);
}
