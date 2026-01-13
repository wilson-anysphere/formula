import type { DrawingTransform, Rect } from "./types";
import { applyTransformVector } from "./transform";

/**
 * Resize handles are drawn in the drawing overlay's screen-space coordinate system
 * (post-zoom, pre-DPR). Keeping these sizes in px ensures the handles remain a
 * consistent on-screen size across zoom levels.
 */
export const RESIZE_HANDLE_SIZE_PX = 8;
export const RESIZE_HANDLE_HIT_SIZE_PX = 10;

export type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export interface ResizeHandleCenter {
  handle: ResizeHandle;
  x: number;
  y: number;
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): boolean {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

export function getResizeHandleCenters(bounds: Rect, transform?: DrawingTransform): ResizeHandleCenter[] {
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const hw = bounds.width / 2;
  const hh = bounds.height / 2;

  const local = [
    { handle: "nw" as const, x: -hw, y: -hh },
    { handle: "n" as const, x: 0, y: -hh },
    { handle: "ne" as const, x: hw, y: -hh },
    { handle: "e" as const, x: hw, y: 0 },
    { handle: "se" as const, x: hw, y: hh },
    { handle: "s" as const, x: 0, y: hh },
    { handle: "sw" as const, x: -hw, y: hh },
    { handle: "w" as const, x: -hw, y: 0 },
  ];

  if (!hasNonIdentityTransform(transform)) {
    return local.map((p) => ({ handle: p.handle, x: cx + p.x, y: cy + p.y }));
  }

  return local.map((p) => {
    const t = applyTransformVector(p.x, p.y, transform!);
    return { handle: p.handle, x: cx + t.x, y: cy + t.y };
  });
}

export function hitTestResizeHandle(
  bounds: Rect,
  x: number,
  y: number,
  transform?: DrawingTransform,
): ResizeHandle | null {
  const size = RESIZE_HANDLE_HIT_SIZE_PX;
  const half = size / 2;
  if (!hasNonIdentityTransform(transform)) {
    const x1 = bounds.x;
    const y1 = bounds.y;
    const x2 = bounds.x + bounds.width;
    const y2 = bounds.y + bounds.height;
    const cx = bounds.x + bounds.width / 2;
    const cy = bounds.y + bounds.height / 2;

    // Test in the same order as `getResizeHandleCenters` for deterministic behavior.
    if (x >= x1 - half && x <= x1 + half && y >= y1 - half && y <= y1 + half) return "nw";
    if (x >= cx - half && x <= cx + half && y >= y1 - half && y <= y1 + half) return "n";
    if (x >= x2 - half && x <= x2 + half && y >= y1 - half && y <= y1 + half) return "ne";
    if (x >= x2 - half && x <= x2 + half && y >= cy - half && y <= cy + half) return "e";
    if (x >= x2 - half && x <= x2 + half && y >= y2 - half && y <= y2 + half) return "se";
    if (x >= cx - half && x <= cx + half && y >= y2 - half && y <= y2 + half) return "s";
    if (x >= x1 - half && x <= x1 + half && y >= y2 - half && y <= y2 + half) return "sw";
    if (x >= x1 - half && x <= x1 + half && y >= cy - half && y <= cy + half) return "w";
    return null;
  }

  for (const c of getResizeHandleCenters(bounds, transform)) {
    if (x >= c.x - half && x <= c.x + half && y >= c.y - half && y <= c.y + half) {
      return c.handle;
    }
  }
  return null;
}

export function cursorForResizeHandle(handle: ResizeHandle, transform?: DrawingTransform): string {
  // Fast path for untransformed objects (the common case).
  if (!hasNonIdentityTransform(transform)) {
    switch (handle) {
      case "nw":
      case "se":
        return "nwse-resize";
      case "ne":
      case "sw":
        return "nesw-resize";
      case "n":
      case "s":
        return "ns-resize";
      case "e":
      case "w":
        return "ew-resize";
    }
  }

  // For rotated/flipped objects, map the local handle axis into screen-space and
  // choose the closest available CSS cursor.
  //
  // We only have four cursor options (horizontal/vertical + two diagonals), so
  // this is an approximation — but it keeps 90° rotations and flips intuitive.
  const axis = (() => {
    switch (handle) {
      case "n":
      case "s":
        return { x: 0, y: 1 };
      case "e":
      case "w":
        return { x: 1, y: 0 };
      case "nw":
      case "se":
        return { x: 1, y: 1 };
      case "ne":
      case "sw":
        return { x: 1, y: -1 };
    }
  })();

  const v = applyTransformVector(axis.x, axis.y, transform!);
  const ax = Math.abs(v.x);
  const ay = Math.abs(v.y);

  if (handle === "n" || handle === "s" || handle === "e" || handle === "w") {
    // Edge handles: choose horizontal vs vertical based on dominant axis.
    return ax >= ay ? "ew-resize" : "ns-resize";
  }

  // Corner handles: choose the diagonal based on the transformed axis orientation.
  return v.x * v.y >= 0 ? "nwse-resize" : "nesw-resize";
}
