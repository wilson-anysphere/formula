import { formatNumber } from "./format.js";
import type { Transform } from "./types.js";

export function transformToSvg(transform: Transform[] | undefined): string | null {
  if (!transform?.length) return null;
  const parts: string[] = [];
  for (const t of transform) {
    switch (t.kind) {
      case "translate":
        parts.push(`translate(${formatNumber(t.x)} ${formatNumber(t.y)})`);
        break;
      case "scale": {
        const sy = t.y ?? t.x;
        parts.push(`scale(${formatNumber(t.x)} ${formatNumber(sy)})`);
        break;
      }
      case "rotate": {
        const deg = (t.radians * 180) / Math.PI;
        if (t.cx == null || t.cy == null) parts.push(`rotate(${formatNumber(deg)})`);
        else parts.push(`rotate(${formatNumber(deg)} ${formatNumber(t.cx)} ${formatNumber(t.cy)})`);
        break;
      }
    }
  }
  return parts.join(" ");
}

export function applyTransformToCanvas(ctx: CanvasRenderingContext2D, transform: Transform[] | undefined): void {
  if (!transform?.length) return;
  for (const t of transform) {
    switch (t.kind) {
      case "translate":
        ctx.translate(t.x, t.y);
        break;
      case "scale":
        ctx.scale(t.x, t.y ?? t.x);
        break;
      case "rotate":
        if (t.cx == null || t.cy == null) {
          ctx.rotate(t.radians);
          break;
        }
        ctx.translate(t.cx, t.cy);
        ctx.rotate(t.radians);
        ctx.translate(-t.cx, -t.cy);
        break;
    }
  }
}

