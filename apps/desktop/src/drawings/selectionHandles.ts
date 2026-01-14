import type { DrawingTransform, Rect } from "./types";
import { DRAWING_HANDLE_SIZE_PX } from "./constants";

/**
 * Resize handles are drawn in the drawing overlay's screen-space coordinate system
 * (post-zoom, pre-DPR). Keeping these sizes in px ensures the handles remain a
 * consistent on-screen size across zoom levels.
 */
export const RESIZE_HANDLE_SIZE_PX = DRAWING_HANDLE_SIZE_PX;

// Rotation handle is drawn separately from resize handles (Excel-style). We keep
// it slightly larger to avoid it being confused with the top midpoint handle.
export const ROTATION_HANDLE_SIZE_PX = 10;
export const ROTATION_HANDLE_HIT_SIZE_PX = 14;
// Distance from the selection outline (top edge midpoint) to the rotation handle.
export const ROTATION_HANDLE_OFFSET_PX = 18;

export type ResizeHandle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";

export interface ResizeHandleCenter {
  handle: ResizeHandle;
  x: number;
  y: number;
}

export interface RotationHandleCenter {
  x: number;
  y: number;
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): transform is DrawingTransform {
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

const RESIZE_HANDLE_ORDER: readonly ResizeHandle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
const RESIZE_HANDLE_SX: readonly number[] = [-1, 0, 1, 1, 1, 0, -1, -1];
const RESIZE_HANDLE_SY: readonly number[] = [-1, -1, -1, 0, 1, 1, 1, 0];
const ROTATION_EDGE_SX: readonly number[] = [0, 1, 0, -1];
const ROTATION_EDGE_SY: readonly number[] = [-1, 0, 1, 0];

function ensureResizeHandleCentersOut(out: ResizeHandleCenter[]): ResizeHandleCenter[] {
  for (let i = 0; i < RESIZE_HANDLE_ORDER.length; i += 1) {
    const handle = RESIZE_HANDLE_ORDER[i]!;
    const existing = out[i];
    if (existing) {
      existing.handle = handle;
      continue;
    }
    out[i] = { handle, x: 0, y: 0 };
  }
  out.length = RESIZE_HANDLE_ORDER.length;
  return out;
}

/**
 * Computes resize handle centers, writing into `out` to avoid allocations.
 *
 * `out` is normalized to always contain 8 entries in the canonical handle order:
 *
 *   nw, n, ne, e, se, s, sw, w
 */
export function getResizeHandleCentersInto(
  bounds: Rect,
  transform: DrawingTransform | undefined,
  out: ResizeHandleCenter[],
): ResizeHandleCenter[] {
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const hw = bounds.width / 2;
  const hh = bounds.height / 2;

  const points = ensureResizeHandleCentersOut(out);

  if (!hasNonIdentityTransform(transform)) {
    for (let i = 0; i < points.length; i += 1) {
      const dx = RESIZE_HANDLE_SX[i]! * hw;
      const dy = RESIZE_HANDLE_SY[i]! * hh;
      points[i]!.x = cx + dx;
      points[i]!.y = cy + dy;
    }
    return points;
  }

  const trig = getTransformTrig(transform);
  const cos = trig.cos;
  const sin = trig.sin;
  const flipH = transform.flipH;
  const flipV = transform.flipV;

  for (let i = 0; i < points.length; i += 1) {
    let x = RESIZE_HANDLE_SX[i]! * hw;
    let y = RESIZE_HANDLE_SY[i]! * hh;
    if (flipH) x = -x;
    if (flipV) y = -y;
    // Forward transform: scale(flip) then rotate(theta).
    const tx = x * cos - y * sin;
    const ty = x * sin + y * cos;
    points[i]!.x = cx + tx;
    points[i]!.y = cy + ty;
  }
  return points;
}

export function getResizeHandleCenters(bounds: Rect, transform?: DrawingTransform): ResizeHandleCenter[] {
  return getResizeHandleCentersInto(bounds, transform, []);
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
  const trig = getTransformTrig(transform);
  const cos = trig.cos;
  const sin = trig.sin;
  const flipH = transform.flipH;
  const flipV = transform.flipV;

  // Test in the same order as `getResizeHandleCenters` for deterministic behavior.
  for (let i = 0; i < RESIZE_HANDLE_ORDER.length; i += 1) {
    const handle = RESIZE_HANDLE_ORDER[i]!;
    let dx = RESIZE_HANDLE_SX[i]! * hw;
    let dy = RESIZE_HANDLE_SY[i]! * hh;
    if (flipH) dx = -dx;
    if (flipV) dy = -dy;
    const tx = dx * cos - dy * sin;
    const ty = dx * sin + dy * cos;
    const hx = cx + tx;
    const hy = cy + ty;
    if (x >= hx - half && x <= hx + half && y >= hy - half && y <= hy + half) return handle;
  }
  return null;
}

export function getRotationHandleCenter(bounds: Rect, transform?: DrawingTransform): RotationHandleCenter {
  const out: RotationHandleCenter = { x: 0, y: 0 };
  return getRotationHandleCenterInto(bounds, transform, out);
}

/**
 * Allocation-free rotation handle center calculation.
 *
 * Writes into `out` and returns it.
 */
export function getRotationHandleCenterInto(
  bounds: Rect,
  transform: DrawingTransform | undefined,
  out: RotationHandleCenter,
): RotationHandleCenter {
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const hw = bounds.width / 2;
  const hh = bounds.height / 2;

  let topX = cx;
  let topY = cy - hh;

  if (hasNonIdentityTransform(transform)) {
    const trig = getTransformTrig(transform);
    const cos = trig.cos;
    const sin = trig.sin;
    const flipH = transform.flipH;
    const flipV = transform.flipV;

    let bestX = 0;
    let bestY = 0;
    let hasBest = false;

    for (let i = 0; i < 4; i += 1) {
      let dx = ROTATION_EDGE_SX[i]! * hw;
      let dy = ROTATION_EDGE_SY[i]! * hh;
      if (flipH) dx = -dx;
      if (flipV) dy = -dy;
      const tx = dx * cos - dy * sin;
      const ty = dx * sin + dy * cos;
      const wx = cx + tx;
      const wy = cy + ty;
      if (!hasBest || wy < bestY) {
        bestX = wx;
        bestY = wy;
        hasBest = true;
      }
    }

    if (hasBest) {
      topX = bestX;
      topY = bestY;
    }
  }

  const vx = topX - cx;
  const vy = topY - cy;
  const len = Math.hypot(vx, vy);
  const ux = len > 0 ? vx / len : 0;
  const uy = len > 0 ? vy / len : -1;

  out.x = topX + ux * ROTATION_HANDLE_OFFSET_PX;
  out.y = topY + uy * ROTATION_HANDLE_OFFSET_PX;
  return out;
}

export function hitTestRotationHandle(
  bounds: Rect,
  x: number,
  y: number,
  transform?: DrawingTransform,
): boolean {
  const cx = bounds.x + bounds.width / 2;
  const cy = bounds.y + bounds.height / 2;
  const hw = bounds.width / 2;
  const hh = bounds.height / 2;

  let topX = cx;
  let topY = cy - hh;

  if (hasNonIdentityTransform(transform)) {
    const trig = getTransformTrig(transform);
    const cos = trig.cos;
    const sin = trig.sin;
    const flipH = transform.flipH;
    const flipV = transform.flipV;

    let bestX = 0;
    let bestY = 0;
    let hasBest = false;

    for (let i = 0; i < 4; i += 1) {
      let dx = ROTATION_EDGE_SX[i]! * hw;
      let dy = ROTATION_EDGE_SY[i]! * hh;
      if (flipH) dx = -dx;
      if (flipV) dy = -dy;
      const tx = dx * cos - dy * sin;
      const ty = dx * sin + dy * cos;
      const wx = cx + tx;
      const wy = cy + ty;
      if (!hasBest || wy < bestY) {
        bestX = wx;
        bestY = wy;
        hasBest = true;
      }
    }
    if (hasBest) {
      topX = bestX;
      topY = bestY;
    }
  }

  const vx = topX - cx;
  const vy = topY - cy;
  const len = Math.hypot(vx, vy);
  const ux = len > 0 ? vx / len : 0;
  const uy = len > 0 ? vy / len : -1;

  const handleX = topX + ux * ROTATION_HANDLE_OFFSET_PX;
  const handleY = topY + uy * ROTATION_HANDLE_OFFSET_PX;
  const half = ROTATION_HANDLE_HIT_SIZE_PX / 2;
  return x >= handleX - half && x <= handleX + half && y >= handleY - half && y <= handleY + half;
}

export function cursorForRotationHandle(active?: boolean): string {
  // There is no cross-browser standard rotation cursor. Use grab/grabbing for an
  // intuitive affordance.
  return active ? "grabbing" : "grab";
}

function cursorForResizeHandleUntransformed(handle: ResizeHandle): string {
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

/**
 * Cursor shape for a resize handle, optionally accounting for a drawing transform.
 *
 * NOTE: This function historically accepted an (ignored) second argument at many
 * call sites. We keep the optional `transform` parameter for compatibility so
 * unit tests and older callers can pass it without needing a separate import.
 */
export function cursorForResizeHandle(handle: ResizeHandle, transform?: DrawingTransform): string {
  if (hasNonIdentityTransform(transform)) return cursorForResizeHandleWithTransform(handle, transform);
  return cursorForResizeHandleUntransformed(handle);
}

export function cursorForResizeHandleWithTransform(handle: ResizeHandle, transform?: DrawingTransform): string {
  // Match the legacy mapping for the common "no transform" case.
  if (!hasNonIdentityTransform(transform)) return cursorForResizeHandleUntransformed(handle);

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

  const trig = getTransformTrig(transform);
  const cos = trig.cos;
  const sin = trig.sin;
  let vx = transform.flipH ? -dx : dx;
  let vy = transform.flipV ? -dy : dy;
  const tx = vx * cos - vy * sin;
  const ty = vx * sin + vy * cos;
  const angleDeg = (Math.atan2(ty, tx) * 180) / Math.PI;
  // Reduce to [0, 180) since cursor direction is symmetric under 180° rotation.
  const normalized = ((angleDeg % 180) + 180) % 180;
  // Snap to nearest 45° (0, 45, 90, 135) modulo 180.
  const rawIndex = Math.round(normalized / 45);
  if (!Number.isFinite(rawIndex)) return cursorForResizeHandleUntransformed(handle);
  const snappedIndex = ((rawIndex % 4) + 4) % 4;
  return SNAP_CURSOR_BY_45_DEG[snappedIndex];
}
