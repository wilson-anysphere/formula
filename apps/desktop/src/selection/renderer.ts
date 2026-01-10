import type { CellCoord, Range, SelectionState } from "./types";

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

const DEFAULT_STYLE: SelectionRenderStyle = {
  fillColor: "rgba(14, 101, 235, 0.12)",
  borderColor: "#0e65eb",
  activeBorderColor: "#0e65eb",
  borderWidth: 2,
  activeBorderWidth: 3,
  fillHandleSize: 8
};

export class SelectionRenderer {
  constructor(private style: SelectionRenderStyle = DEFAULT_STYLE) {}

  render(ctx: CanvasRenderingContext2D, selection: SelectionState, metrics: GridMetrics): void {
    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();

    // We draw in CSS pixels; the caller should already have adjusted for DPR.
    this.renderFill(ctx, selection.ranges, metrics);
    this.renderBorders(ctx, selection.ranges, metrics);
    this.renderActiveCell(ctx, selection.active, metrics, selection.type);
  }

  private renderFill(ctx: CanvasRenderingContext2D, ranges: Range[], metrics: GridMetrics) {
    ctx.save();
    ctx.fillStyle = this.style.fillColor;
    for (const range of ranges) {
      const rect = this.rangeToRect(range, metrics);
      if (!rect) continue;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    ctx.restore();
  }

  private renderBorders(ctx: CanvasRenderingContext2D, ranges: Range[], metrics: GridMetrics) {
    ctx.save();
    ctx.strokeStyle = this.style.borderColor;
    ctx.lineWidth = this.style.borderWidth;
    for (const range of ranges) {
      const rect = this.rangeToRect(range, metrics);
      if (!rect) continue;
      // Sub-pixel alignment for crisp borders.
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
    }
    ctx.restore();
  }

  private renderActiveCell(ctx: CanvasRenderingContext2D, cell: CellCoord, metrics: GridMetrics, selectionType: SelectionState["type"]) {
    const rect = metrics.getCellRect(cell);
    if (!rect) return;

    ctx.save();
    ctx.strokeStyle = this.style.activeBorderColor;
    ctx.lineWidth = this.style.activeBorderWidth;
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);

    if (selectionType !== "row" && selectionType !== "column" && selectionType !== "all") {
      const size = this.style.fillHandleSize;
      ctx.fillStyle = this.style.activeBorderColor;
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
