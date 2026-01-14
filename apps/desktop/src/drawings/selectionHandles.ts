import type { DrawingTransform, Rect } from "./types";
import { applyTransformVector } from "./transform";
import { DRAWING_HANDLE_SIZE_PX } from "./constants";

/**
 * Resize handles are drawn in the drawing overlay's screen-space coordinate system
 * (post-zoom, pre-DPR). Keeping these sizes in px ensures the handles remain a
 * consistent on-screen size across zoom levels.
 */
export const RESIZE_HANDLE_SIZE_PX = DRAWING_HANDLE_SIZE_PX;

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

type CachedTrig = { rotationDeg: number; cos: number; sin: number };

const trigCache = new WeakMap<DrawingTransform, CachedTrig>();

function getTransformTrig(transform: DrawingTransform): CachedTrig {
  const cached = trigCache.get(transform);
  const rot = transform.rotationDeg;
  if (cached && cached.rotationDeg === rot) return cached;
  const radians = (rot * Math.PI) / 180;
  const next: CachedTrig = { rotationDeg: rot, cos: Math.cos(radians), sin: Math.sin(radians) };
  trigCache.set(transform, next);
  return next;
}

function applyTransformVectorFast(dx: number, dy: number, transform: DrawingTransform, trig: CachedTrig): { x: number; y: number } {
  let x = dx;
  let y = dy;
  if (transform.flipH) x = -x;
  if (transform.flipV) y = -y;
  return {
    x: x * trig.cos - y * trig.sin,
    y: x * trig.sin + y * trig.cos,
  };
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

  const trig = getTransformTrig(transform!);
  return local.map((p) => {
    const t = applyTransformVectorFast(p.x, p.y, transform!, trig);
    return { handle: p.handle, x: cx + t.x, y: cy + t.y };
  });
}

export function hitTestResizeHandle(
  bounds: Rect,
  x: number,
  y: number,
  transform?: DrawingTransform,
): ResizeHandle | null {
  // Keep the interactive hit target in sync with the rendered handle square so
  // the visible geometry matches the resize affordance.
  const size = RESIZE_HANDLE_SIZE_PX;
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

  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const hw = bounds.width / 2;
  const hh = bounds.height / 2;
  const trig = getTransformTrig(transform!);

  const test = (handle: ResizeHandle, dx: number, dy: number): ResizeHandle | null => {
    const t = applyTransformVectorFast(dx, dy, transform!, trig);
    const hx = cx + t.x;
    const hy = cy + t.y;
    if (x >= hx - half && x <= hx + half && y >= hy - half && y <= hy + half) return handle;
    return null;
  };

  // Test in the same order as `getResizeHandleCenters` for deterministic behavior.
  return (
    test("nw", -hw, -hh) ??
    test("n", 0, -hh) ??
    test("ne", hw, -hh) ??
    test("e", hw, 0) ??
    test("se", hw, hh) ??
    test("s", 0, hh) ??
    test("sw", -hw, hh) ??
    test("w", -hw, 0)
  );
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

const SNAP_CURSOR_BY_45_DEG: readonly [string, string, string, string] = [
  "ew-resize",
  "nwse-resize",
  "ns-resize",
  "nesw-resize",
];

export function cursorForResizeHandleWithTransform(handle: ResizeHandle, transform?: DrawingTransform): string {
  // Match the legacy mapping for the common "no transform" case.
  if (!hasNonIdentityTransform(transform)) return cursorForResizeHandle(handle);

  let dx = 0;
  let dy = 0;
  switch (handle) {
    case "nw":
      dx = -1;
      dy = -1;
      break;
    case "n":
      dx = 0;
      dy = -1;
      break;
    case "ne":
      dx = 1;
      dy = -1;
      break;
    case "e":
      dx = 1;
      dy = 0;
      break;
    case "se":
      dx = 1;
      dy = 1;
      break;
    case "s":
      dx = 0;
      dy = 1;
      break;
    case "sw":
      dx = -1;
      dy = 1;
      break;
    case "w":
      dx = -1;
      dy = 0;
      break;
  }

  const v = applyTransformVector(dx, dy, transform!);
  const angleDeg = (Math.atan2(v.y, v.x) * 180) / Math.PI;
  // Reduce to [0, 180) since cursor direction is symmetric under 180° rotation.
  const normalized = ((angleDeg % 180) + 180) % 180;
  // Snap to nearest 45° (0, 45, 90, 135) modulo 180.
  const rawIndex = Math.round(normalized / 45);
  if (!Number.isFinite(rawIndex)) return cursorForResizeHandle(handle);
  const snappedIndex = ((rawIndex % 4) + 4) % 4;
  return SNAP_CURSOR_BY_45_DEG[snappedIndex];
}
