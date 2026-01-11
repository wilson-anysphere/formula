import type { CellCoord, Range, SelectionState } from "./types";

import { resolveCssVar } from "../theme/cssVars.js";

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface GridMetrics {
  getCellRect(cell: CellCoord): Rect | null;
  /**
   * Visible row indices (0-based) in the current viewport.
   */
  visibleRows: readonly number[];
  /**
   * Visible column indices (0-based) in the current viewport.
   */
  visibleCols: readonly number[];
}

export interface SelectionRenderStyle {
  fillColor: string;
  borderColor: string;
  activeBorderColor: string;
  borderWidth: number;
  activeBorderWidth: number;
  fillHandleSize: number;
}

export type SelectionRangeRenderInfo = {
  range: Range;
  rect: Rect;
  edges: { top: boolean; right: boolean; bottom: boolean; left: boolean };
};

export type SelectionRenderDebugInfo = {
  ranges: SelectionRangeRenderInfo[];
  activeCellRect: Rect | null;
};

function defaultStyleFromTheme(): SelectionRenderStyle {
  return {
    fillColor: resolveCssVar("--selection-bg", { fallback: "transparent" }),
    borderColor: resolveCssVar("--selection-border", { fallback: "transparent" }),
    activeBorderColor: resolveCssVar("--selection-border", { fallback: "transparent" }),
    borderWidth: 2,
    activeBorderWidth: 3,
    fillHandleSize: 8,
  };
}

export class SelectionRenderer {
  constructor(private style: SelectionRenderStyle | null = null) {}

  private lastDebug: SelectionRenderDebugInfo | null = null;

  getLastDebug(): SelectionRenderDebugInfo | null {
    return this.lastDebug;
  }

  render(ctx: CanvasRenderingContext2D, selection: SelectionState, metrics: GridMetrics): void {
    const style = this.style ?? defaultStyleFromTheme();

    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();

    const visibleRanges = this.computeVisibleRanges(selection.ranges, metrics);

    this.lastDebug = {
      ranges: visibleRanges,
      activeCellRect: metrics.getCellRect(selection.active),
    };

    // We draw in CSS pixels; the caller should already have adjusted for DPR.
    this.renderFill(ctx, visibleRanges, style);
    this.renderBorders(ctx, visibleRanges, style);
    this.renderActiveCell(ctx, selection.active, metrics, selection.type, style);
  }

  private renderFill(ctx: CanvasRenderingContext2D, ranges: SelectionRangeRenderInfo[], style: SelectionRenderStyle) {
    ctx.save();
    ctx.fillStyle = style.fillColor;
    for (const range of ranges) {
      const rect = range.rect;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    ctx.restore();
  }

  private renderBorders(ctx: CanvasRenderingContext2D, ranges: SelectionRangeRenderInfo[], style: SelectionRenderStyle) {
    ctx.save();
    ctx.strokeStyle = style.borderColor;
    ctx.lineWidth = style.borderWidth;
    for (const range of ranges) {
      this.strokeVisibleRange(ctx, range);
    }
    ctx.restore();
  }

  private renderActiveCell(
    ctx: CanvasRenderingContext2D,
    cell: CellCoord,
    metrics: GridMetrics,
    selectionType: SelectionState["type"],
    style: SelectionRenderStyle,
  ) {
    const rect = metrics.getCellRect(cell);
    if (!rect) return;

    ctx.save();
    ctx.strokeStyle = style.activeBorderColor;
    ctx.lineWidth = style.activeBorderWidth;
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);

    if (selectionType !== "row" && selectionType !== "column" && selectionType !== "all") {
      const size = style.fillHandleSize;
      ctx.fillStyle = style.activeBorderColor;
      ctx.fillRect(rect.x + rect.width - size / 2, rect.y + rect.height - size / 2, size, size);
    }

    ctx.restore();
  }

  private computeVisibleRanges(ranges: Range[], metrics: GridMetrics): SelectionRangeRenderInfo[] {
    const out: SelectionRangeRenderInfo[] = [];
    for (const range of ranges) {
      const info = this.rangeToVisibleRange(range, metrics);
      if (!info) continue;
      out.push(info);
    }
    return out;
  }

  private rangeToVisibleRange(range: Range, metrics: GridMetrics): SelectionRangeRenderInfo | null {
    const visibleStartRow = firstVisibleIndex(metrics.visibleRows, range.startRow, range.endRow);
    const visibleEndRow = lastVisibleIndex(metrics.visibleRows, range.startRow, range.endRow);
    const visibleStartCol = firstVisibleIndex(metrics.visibleCols, range.startCol, range.endCol);
    const visibleEndCol = lastVisibleIndex(metrics.visibleCols, range.startCol, range.endCol);
    if (visibleStartRow == null || visibleEndRow == null || visibleStartCol == null || visibleEndCol == null) {
      return null;
    }

    const start = metrics.getCellRect({ row: visibleStartRow, col: visibleStartCol });
    const end = metrics.getCellRect({ row: visibleEndRow, col: visibleEndCol });
    if (!start || !end) return null;

    const x = start.x;
    const y = start.y;
    const width = end.x + end.width - start.x;
    const height = end.y + end.height - start.y;
    if (width <= 0 || height <= 0) return null;

    const edges = {
      top: metrics.getCellRect({ row: range.startRow, col: visibleStartCol }) != null,
      bottom: metrics.getCellRect({ row: range.endRow, col: visibleStartCol }) != null,
      left: metrics.getCellRect({ row: visibleStartRow, col: range.startCol }) != null,
      right: metrics.getCellRect({ row: visibleStartRow, col: range.endCol }) != null,
    };

    return { range, rect: { x, y, width, height }, edges };
  }

  private strokeVisibleRange(ctx: CanvasRenderingContext2D, range: SelectionRangeRenderInfo): void {
    const { rect, edges } = range;

    // Sub-pixel alignment for crisp borders.
    const xLeft = rect.x + 0.5;
    const xRight = rect.x + rect.width - 0.5;
    const yTop = rect.y + 0.5;
    const yBottom = rect.y + rect.height - 0.5;

    // Fast path: fully visible range => draw a single rectangle.
    if (edges.top && edges.right && edges.bottom && edges.left) {
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
      return;
    }

    // Draw only edges whose boundaries are visible in the viewport. This avoids
    // drawing a misleading "clamped" border when the selection extends offscreen.
    ctx.beginPath();
    if (edges.top) {
      ctx.moveTo(xLeft, yTop);
      ctx.lineTo(xRight, yTop);
    }
    if (edges.bottom) {
      ctx.moveTo(xLeft, yBottom);
      ctx.lineTo(xRight, yBottom);
    }
    if (edges.left) {
      ctx.moveTo(xLeft, yTop);
      ctx.lineTo(xLeft, yBottom);
    }
    if (edges.right) {
      ctx.moveTo(xRight, yTop);
      ctx.lineTo(xRight, yBottom);
    }
    ctx.stroke();
  }
}

function firstVisibleIndex(values: readonly number[], start: number, end: number): number | null {
  for (const value of values) {
    if (value < start) continue;
    if (value > end) break;
    return value;
  }
  return null;
}

function lastVisibleIndex(values: readonly number[], start: number, end: number): number | null {
  for (let i = values.length - 1; i >= 0; i -= 1) {
    const value = values[i]!;
    if (value > end) continue;
    if (value < start) break;
    return value;
  }
  return null;
}
