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

    const strokeWidth = this.strokeWidth;
    const halfStroke = strokeWidth / 2;

    ctx.save();
    ctx.lineWidth = strokeWidth;

    ctx.fillStyle = this.precedentFill;
    ctx.strokeStyle = this.precedentStroke;
    for (const addr of highlights.precedents) {
      const parsed = parseCellAddress(addr);
      if (!parsed) continue;
      const rect = getCellRect(parsed.row, parsed.col);
      if (!rect) continue;
      ctx.globalAlpha = this.fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = this.strokeAlpha;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);
    }

    ctx.fillStyle = this.dependentFill;
    ctx.strokeStyle = this.dependentStroke;
    for (const addr of highlights.dependents) {
      const parsed = parseCellAddress(addr);
      if (!parsed) continue;
      const rect = getCellRect(parsed.row, parsed.col);
      if (!rect) continue;
      ctx.globalAlpha = this.fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = this.strokeAlpha;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);
    }

    ctx.restore();
  }
}
