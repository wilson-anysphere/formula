import { parseCellAddress } from "./address";
import type { CellAddress } from "./address";
import { resolveCssVar } from "../../theme/cssVars.js";

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface AuditingHighlights {
  precedents: Iterable<CellAddress>;
  dependents: Iterable<CellAddress>;
}

export interface AuditingOverlayOptions {
  getCellRect: (row: number, col: number) => Rect | null;
}

function withAlpha(color: string, alpha: number) {
  return { color, alpha };
}

function drawCellHighlight(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  options: { fill?: { color: string; alpha: number }; stroke?: { color: string; alpha: number }; strokeWidth?: number },
) {
  const { fill = null, stroke = null, strokeWidth = 2 } = options;

  ctx.save();

  if (fill) {
    ctx.fillStyle = fill.color;
    ctx.globalAlpha = fill.alpha;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
  }

  if (stroke) {
    ctx.strokeStyle = stroke.color;
    ctx.globalAlpha = stroke.alpha;
    ctx.lineWidth = strokeWidth;
    ctx.strokeRect(
      rect.x + strokeWidth / 2,
      rect.y + strokeWidth / 2,
      rect.width - strokeWidth,
      rect.height - strokeWidth,
    );
  }

  ctx.restore();
}

export class AuditingOverlayRenderer {
  precedentFill: string;
  precedentStroke: string;
  dependentFill: string;
  dependentStroke: string;
  fillAlpha: number;
  strokeAlpha: number;
  strokeWidth: number;

  constructor(options?: Partial<AuditingOverlayRenderer>) {
    // Default to theme tokens so overlays match the current UI theme.
    this.precedentFill = options?.precedentFill ?? resolveCssVar("--accent", { fallback: "CanvasText" });
    this.precedentStroke = options?.precedentStroke ?? resolveCssVar("--accent", { fallback: "CanvasText" });
    this.dependentFill = options?.dependentFill ?? resolveCssVar("--success", { fallback: "CanvasText" });
    this.dependentStroke = options?.dependentStroke ?? resolveCssVar("--success", { fallback: "CanvasText" });
    this.fillAlpha = options?.fillAlpha ?? 0.18;
    this.strokeAlpha = options?.strokeAlpha ?? 0.85;
    this.strokeWidth = options?.strokeWidth ?? 2;
  }

  clear(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();
  }

  render(ctx: CanvasRenderingContext2D, highlights: AuditingHighlights, options: AuditingOverlayOptions) {
    const { getCellRect } = options ?? {};
    if (typeof getCellRect !== "function") return;

    for (const addr of highlights.precedents) {
      const parsed = parseCellAddress(addr);
      if (!parsed) continue;
      const rect = getCellRect(parsed.row, parsed.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, {
        fill: withAlpha(this.precedentFill, this.fillAlpha),
        stroke: withAlpha(this.precedentStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
      });
    }

    for (const addr of highlights.dependents) {
      const parsed = parseCellAddress(addr);
      if (!parsed) continue;
      const rect = getCellRect(parsed.row, parsed.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, {
        fill: withAlpha(this.dependentFill, this.fillAlpha),
        stroke: withAlpha(this.dependentStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
      });
    }
  }
}
