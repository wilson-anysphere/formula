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
}

export interface SelectionRenderStyle {
  fillColor: string;
  borderColor: string;
  activeBorderColor: string;
  borderWidth: number;
  activeBorderWidth: number;
  fillHandleSize: number;
}

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

  render(ctx: CanvasRenderingContext2D, selection: SelectionState, metrics: GridMetrics): void {
    const style = this.style ?? defaultStyleFromTheme();

    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();

    // We draw in CSS pixels; the caller should already have adjusted for DPR.
    this.renderFill(ctx, selection.ranges, metrics, style);
    this.renderBorders(ctx, selection.ranges, metrics, style);
    this.renderActiveCell(ctx, selection.active, metrics, selection.type, style);
  }

  private renderFill(ctx: CanvasRenderingContext2D, ranges: Range[], metrics: GridMetrics, style: SelectionRenderStyle) {
    ctx.save();
    ctx.fillStyle = style.fillColor;
    for (const range of ranges) {
      const rect = this.rangeToRect(range, metrics);
      if (!rect) continue;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    ctx.restore();
  }

  private renderBorders(ctx: CanvasRenderingContext2D, ranges: Range[], metrics: GridMetrics, style: SelectionRenderStyle) {
    ctx.save();
    ctx.strokeStyle = style.borderColor;
    ctx.lineWidth = style.borderWidth;
    for (const range of ranges) {
      const rect = this.rangeToRect(range, metrics);
      if (!rect) continue;
      // Sub-pixel alignment for crisp borders.
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
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

  private rangeToRect(range: Range, metrics: GridMetrics): Rect | null {
    const start = metrics.getCellRect({ row: range.startRow, col: range.startCol });
    const end = metrics.getCellRect({ row: range.endRow, col: range.endCol });
    if (!start || !end) return null;

    const x = start.x;
    const y = start.y;
    const width = end.x + end.width - start.x;
    const height = end.y + end.height - start.y;
    if (width <= 0 || height <= 0) return null;
    return { x, y, width, height };
  }
}
