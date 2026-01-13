import type { DrawingObject, DrawingTransform, Rect } from "./types";
import { anchorToRectPx } from "./overlay";
import type { GridGeometry, Viewport } from "./overlay";
import { applyTransformVector, inverseTransformVector } from "./transform";

export interface HitTestResult {
  object: DrawingObject;
  /**
   * Screen-space bounds (px) for the object's *untransformed* anchor rectangle.
   *
   * Callers that want oriented handles/outlines should combine this with
   * `object.transform`.
   */
  bounds: Rect;
}

const EMPTY_LIST: number[] = [];
const CELL_SCRATCH = { row: 0, col: 0 };

export interface HitTestIndex {
  /**
   * Objects sorted by zOrder descending (top-most first).
   *
   * The index stores object ordering separately so repeated pointer-move hit tests
   * don't need to allocate/sort arrays.
   */
  ordered: DrawingObject[];
  /**
   * Untransformed sheet-space bounds (px) for each entry in `ordered`.
   *
   * This corresponds to the anchor rectangle before applying DrawingML
   * rotation/flip transforms.
   */
  bounds: Rect[];
  /**
   * Sheet-space AABBs (px) used for spatial bucketing.
   *
   * For transformed objects, this expands `bounds[i]` to the axis-aligned bounding
   * box of the rotated/flipped rectangle, so hit tests can find objects even when
   * their visible geometry extends beyond the untransformed anchor rectangle.
   */
  aabbs: Rect[];
  /**
   * Spatial bins keyed by bucket X then bucket Y.
   *
   * Each bucket stores indices into `ordered` / `bounds` in zOrder-desc order.
   */
  buckets: Map<number, Map<number, number[]>>;
  /** Indices of objects that were too large to bucket efficiently. */
  global: number[];
  bucketSizePx: number;
  /** Geometry used to derive sheet-space bounds (also used for frozen-pane layout fallbacks). */
  geom: GridGeometry;
  /** Object id -> index into `ordered` / `bounds` (useful for selection/cursor logic). */
  byId: Map<number, number>;
}

function clampNumber(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

export function buildHitTestIndex(
  objects: readonly DrawingObject[],
  geom: GridGeometry,
  opts?: {
    /**
     * Spatial bucket size (px). Larger buckets reduce index build cost; smaller
     * buckets reduce per-hit candidate checks.
     */
    bucketSizePx?: number;
    /**
     * Maximum number of buckets an object can occupy before being placed into
     * the `global` list instead of inserted into every bucket.
     */
    maxBucketsPerObject?: number;
  },
): HitTestIndex {
  const bucketSizePx = Math.max(1, Math.floor(opts?.bucketSizePx ?? 256));
  const maxBucketsPerObject = Math.max(1, Math.floor(opts?.maxBucketsPerObject ?? 256));

  // Walk from top to bottom (highest zOrder first).
  const ordered = [...objects].sort((a, b) => b.zOrder - a.zOrder);
  const bounds: Rect[] = new Array(ordered.length);
  const aabbs: Rect[] = new Array(ordered.length);
  const buckets: Map<number, Map<number, number[]>> = new Map();
  const global: number[] = [];
  const byId = new Map<number, number>();

  for (let i = 0; i < ordered.length; i += 1) {
    const obj = ordered[i]!;
    byId.set(obj.id, i);
    const rect = anchorToRectPx(obj.anchor, geom);
    bounds[i] = rect;
    const aabb = hasNonIdentityTransform(obj.transform) ? rectToAabb(rect, obj.transform!) : rect;
    aabbs[i] = aabb;

    const x1 = aabb.x;
    const y1 = aabb.y;
    const x2 = aabb.x + aabb.width;
    const y2 = aabb.y + aabb.height;

    const minBx = Math.floor(x1 / bucketSizePx);
    const maxBx = Math.floor(x2 / bucketSizePx);
    const minBy = Math.floor(y1 / bucketSizePx);
    const maxBy = Math.floor(y2 / bucketSizePx);

    const bucketsWide = maxBx - minBx + 1;
    const bucketsHigh = maxBy - minBy + 1;
    const bucketCount = bucketsWide * bucketsHigh;

    // Very large objects can span a huge number of buckets; treat them as global candidates
    // to keep index build bounded.
    if (!Number.isFinite(bucketCount) || bucketCount > maxBucketsPerObject) {
      global.push(i);
      continue;
    }

    for (let bx = minBx; bx <= maxBx; bx += 1) {
      let col = buckets.get(bx);
      if (!col) {
        col = new Map();
        buckets.set(bx, col);
      }
      for (let by = minBy; by <= maxBy; by += 1) {
        let list = col.get(by);
        if (!list) {
          list = [];
          col.set(by, list);
        }
        list.push(i);
      }
    }
  }

  return { ordered, bounds, aabbs, buckets, global, bucketSizePx, geom, byId };
}

function hitTestCandidateIndex(
  index: HitTestIndex,
  sheetX: number,
  sheetY: number,
  frozenRows: number,
  frozenCols: number,
  inFrozenRows: boolean,
  inFrozenCols: boolean,
): number | null {
  const bx = Math.floor(sheetX / index.bucketSizePx);
  const by = Math.floor(sheetY / index.bucketSizePx);

  const bucket = index.buckets.get(bx)?.get(by);
  const bucketList = bucket ?? EMPTY_LIST;
  const globalList = index.global.length > 0 ? index.global : EMPTY_LIST;

  const hasFrozenPanes = frozenRows !== 0 || frozenCols !== 0;

  let i = 0;
  let j = 0;
  let last = -1;
  while (i < bucketList.length || j < globalList.length) {
    const next =
      j >= globalList.length || (i < bucketList.length && bucketList[i]! <= globalList[j]!)
        ? bucketList[i++]!
        : globalList[j++]!;
    if (next === last) continue;
    last = next;

    const obj = index.ordered[next]!;
    const rect = index.bounds[next]!;
    const aabb = index.aabbs[next]!;
    if (hasFrozenPanes) {
      const anchor = obj.anchor;
      const objInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < frozenRows;
      const objInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < frozenCols;
      // Excel-like pane routing: each drawing belongs to exactly one quadrant, so pointer
      // hits are constrained to the quadrant under the cursor.
      if (objInFrozenRows !== inFrozenRows || objInFrozenCols !== inFrozenCols) continue;
    }

    if (!pointInRect(sheetX, sheetY, aabb)) continue;

    const hit = hasNonIdentityTransform(obj.transform)
      ? pointInTransformedRect(sheetX, sheetY, rect, obj.transform!)
      : true;

    if (hit) {
      return next;
    }
  }

  return null;
}

export function hitTestDrawings(
  index: HitTestIndex,
  viewport: Viewport,
  x: number,
  y: number,
  geom: GridGeometry = index.geom,
): HitTestResult | null {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;

  // Ignore pointer events over the header area; drawings are rendered under headers.
  if (x < headerOffsetX || y < headerOffsetY) return null;

  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;
  let frozenBoundaryX = headerOffsetX;
  let frozenBoundaryY = headerOffsetY;

  if (frozenCols > 0) {
    let raw = viewport.frozenWidthPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        CELL_SCRATCH.row = 0;
        CELL_SCRATCH.col = frozenCols;
        derived = geom.cellOriginPx(CELL_SCRATCH).x;
      } catch {
        derived = 0;
      }
      raw = headerOffsetX + derived;
    }
    frozenBoundaryX = clampNumber(raw as number, headerOffsetX, viewport.width);
  }

  if (frozenRows > 0) {
    let raw = viewport.frozenHeightPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        CELL_SCRATCH.row = frozenRows;
        CELL_SCRATCH.col = 0;
        derived = geom.cellOriginPx(CELL_SCRATCH).y;
      } catch {
        derived = 0;
      }
      raw = headerOffsetY + derived;
    }
    frozenBoundaryY = clampNumber(raw as number, headerOffsetY, viewport.height);
  }

  const inFrozenCols = frozenCols > 0 && x < frozenBoundaryX;
  const inFrozenRows = frozenRows > 0 && y < frozenBoundaryY;

  // Convert from screen-space to sheet-space using the same frozen-pane scroll semantics
  // as `DrawingOverlay.render()`.
  const scrollX = inFrozenCols ? 0 : viewport.scrollX;
  const scrollY = inFrozenRows ? 0 : viewport.scrollY;
  const sheetX = x - headerOffsetX + scrollX;
  const sheetY = y - headerOffsetY + scrollY;

  const hitIndex = hitTestCandidateIndex(index, sheetX, sheetY, frozenRows, frozenCols, inFrozenRows, inFrozenCols);
  if (hitIndex == null) return null;

  const rect = index.bounds[hitIndex]!;
  const screen = {
    x: rect.x - scrollX + headerOffsetX,
    y: rect.y - scrollY + headerOffsetY,
    width: rect.width,
    height: rect.height,
  };
  return { object: index.ordered[hitIndex]!, bounds: screen };
}

export function hitTestDrawingsObject(
  index: HitTestIndex,
  viewport: Viewport,
  x: number,
  y: number,
  geom: GridGeometry = index.geom,
): DrawingObject | null {
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;

  if (x < headerOffsetX || y < headerOffsetY) return null;

  const frozenRows = Number.isFinite(viewport.frozenRows) ? Math.max(0, Math.trunc(viewport.frozenRows!)) : 0;
  const frozenCols = Number.isFinite(viewport.frozenCols) ? Math.max(0, Math.trunc(viewport.frozenCols!)) : 0;

  let frozenBoundaryX = headerOffsetX;
  let frozenBoundaryY = headerOffsetY;

  if (frozenCols > 0) {
    let raw = viewport.frozenWidthPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        CELL_SCRATCH.row = 0;
        CELL_SCRATCH.col = frozenCols;
        derived = geom.cellOriginPx(CELL_SCRATCH).x;
      } catch {
        derived = 0;
      }
      raw = headerOffsetX + derived;
    }
    frozenBoundaryX = clampNumber(raw as number, headerOffsetX, viewport.width);
  }

  if (frozenRows > 0) {
    let raw = viewport.frozenHeightPx;
    if (!Number.isFinite(raw)) {
      let derived = 0;
      try {
        CELL_SCRATCH.row = frozenRows;
        CELL_SCRATCH.col = 0;
        derived = geom.cellOriginPx(CELL_SCRATCH).y;
      } catch {
        derived = 0;
      }
      raw = headerOffsetY + derived;
    }
    frozenBoundaryY = clampNumber(raw as number, headerOffsetY, viewport.height);
  }

  const inFrozenCols = frozenCols > 0 && x < frozenBoundaryX;
  const inFrozenRows = frozenRows > 0 && y < frozenBoundaryY;

  const scrollX = inFrozenCols ? 0 : viewport.scrollX;
  const scrollY = inFrozenRows ? 0 : viewport.scrollY;
  const sheetX = x - headerOffsetX + scrollX;
  const sheetY = y - headerOffsetY + scrollY;

  const hitIndex = hitTestCandidateIndex(index, sheetX, sheetY, frozenRows, frozenCols, inFrozenRows, inFrozenCols);
  return hitIndex == null ? null : index.ordered[hitIndex]!;
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): boolean {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

function pointInTransformedRect(x: number, y: number, rect: Rect, transform: DrawingTransform): boolean {
  if (!(rect.width > 0 && rect.height > 0)) return false;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const local = inverseTransformVector(x - cx, y - cy, transform);
  const hw = rect.width / 2;
  const hh = rect.height / 2;
  return local.x >= -hw && local.x <= hw && local.y >= -hh && local.y <= hh;
}

function rectToAabb(rect: Rect, transform: DrawingTransform): Rect {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;

  const corners = [
    applyTransformVector(-hw, -hh, transform),
    applyTransformVector(hw, -hh, transform),
    applyTransformVector(hw, hh, transform),
    applyTransformVector(-hw, hh, transform),
  ];

  let minX = cx + corners[0]!.x;
  let maxX = cx + corners[0]!.x;
  let minY = cy + corners[0]!.y;
  let maxY = cy + corners[0]!.y;

  for (let i = 1; i < corners.length; i += 1) {
    const p = corners[i]!;
    const x = cx + p.x;
    const y = cy + p.y;
    if (x < minX) minX = x;
    if (x > maxX) maxX = x;
    if (y < minY) minY = y;
    if (y > maxY) maxY = y;
  }

  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}
