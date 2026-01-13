import type { Rect } from "./types";

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

export function getResizeHandleCenters(bounds: Rect): ResizeHandleCenter[] {
  const x1 = bounds.x;
  const y1 = bounds.y;
  const x2 = bounds.x + bounds.width;
  const y2 = bounds.y + bounds.height;
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;

  return [
    { handle: "nw", x: x1, y: y1 },
    { handle: "n", x: cx, y: y1 },
    { handle: "ne", x: x2, y: y1 },
    { handle: "e", x: x2, y: cy },
    { handle: "se", x: x2, y: y2 },
    { handle: "s", x: cx, y: y2 },
    { handle: "sw", x: x1, y: y2 },
    { handle: "w", x: x1, y: cy },
  ];
}

export function hitTestResizeHandle(bounds: Rect, x: number, y: number): ResizeHandle | null {
  const size = RESIZE_HANDLE_HIT_SIZE_PX;
  const half = size / 2;
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

export function cursorForResizeHandle(handle: ResizeHandle): string {
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
