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

export function computeFillPreview(sourceRange: CellRange, pointerCell: { row: number; col: number }): FillDragPreview | null {
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
    const unionRange: CellRange = {
      startRow: Math.min(srcTop, row),
      endRow: Math.max(srcBottomExclusive, row + 1),
      startCol: sourceRange.startCol,
      endCol: sourceRange.endCol
    };

    const targetRange: CellRange | null =
      row >= srcBottomExclusive
        ? {
            startRow: srcBottomExclusive,
            endRow: row + 1,
            startCol: sourceRange.startCol,
            endCol: sourceRange.endCol
          }
        : row < srcTop
          ? {
              startRow: row,
              endRow: srcTop,
              startCol: sourceRange.startCol,
              endCol: sourceRange.endCol
            }
          : null;

    return targetRange ? { axis, targetRange, unionRange } : null;
  }

  const unionRange: CellRange = {
    startRow: sourceRange.startRow,
    endRow: sourceRange.endRow,
    startCol: Math.min(srcLeft, col),
    endCol: Math.max(srcRightExclusive, col + 1)
  };

  const targetRange: CellRange | null =
    col >= srcRightExclusive
      ? {
          startRow: sourceRange.startRow,
          endRow: sourceRange.endRow,
          startCol: srcRightExclusive,
          endCol: col + 1
        }
      : col < srcLeft
        ? {
            startRow: sourceRange.startRow,
            endRow: sourceRange.endRow,
            startCol: col,
            endCol: srcLeft
          }
        : null;

  return targetRange ? { axis, targetRange, unionRange } : null;
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
