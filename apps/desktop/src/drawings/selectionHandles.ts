import type { DrawingTransform, Rect } from "./types";

/**
 * Resize handles are drawn in the drawing overlay's screen-space coordinate system
 * (post-zoom, pre-DPR). Keeping these sizes in px ensures the handles remain a
 * consistent on-screen size across zoom levels.
 */
export const RESIZE_HANDLE_SIZE_PX = 8;

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
  let axisX = 0;
  let axisY = 0;
  switch (handle) {
    case "n":
    case "s":
      axisX = 0;
      axisY = 1;
      break;
    case "e":
    case "w":
      axisX = 1;
      axisY = 0;
      break;
    case "nw":
    case "se":
      axisX = 1;
      axisY = 1;
      break;
    case "ne":
    case "sw":
      axisX = 1;
      axisY = -1;
      break;
  }

  const trig = getTransformTrig(transform!);
  const v = applyTransformVectorFast(axisX, axisY, transform!, trig);
  const ax = Math.abs(v.x);
  const ay = Math.abs(v.y);

  if (handle === "n" || handle === "s" || handle === "e" || handle === "w") {
    // Edge handles: choose horizontal vs vertical based on dominant axis.
    return ax >= ay ? "ew-resize" : "ns-resize";
  }

  // Corner handles: choose the diagonal based on the transformed axis orientation.
  return v.x * v.y >= 0 ? "nwse-resize" : "nesw-resize";
}
