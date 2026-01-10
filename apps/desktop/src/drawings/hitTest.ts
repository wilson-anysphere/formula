import type { DrawingObject, Rect } from "./types";
import { anchorToRectPx } from "./overlay";
import type { GridGeometry, Viewport } from "./overlay";

export interface HitTestResult {
  object: DrawingObject;
  bounds: Rect;
}

export function hitTestDrawings(
  objects: DrawingObject[],
  viewport: Viewport,
  geom: GridGeometry,
  x: number,
  y: number,
): HitTestResult | null {
  // Walk from top to bottom (highest zOrder first).
  const ordered = [...objects].sort((a, b) => b.zOrder - a.zOrder);
  for (const obj of ordered) {
    const rect = anchorToRectPx(obj.anchor, geom);
    const screen = {
      x: rect.x - viewport.scrollX,
      y: rect.y - viewport.scrollY,
      width: rect.width,
      height: rect.height,
    };
    if (pointInRect(x, y, screen)) {
      return { object: obj, bounds: screen };
    }
  }
  return null;
}

function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

