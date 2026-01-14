import type { DrawingObject, DrawingTransform, Rect } from "./types";
import { anchorToRectPx, effectiveScrollForAnchor } from "./overlay";
import type { GridGeometry, Viewport } from "./overlay";

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

export function drawingObjectToViewportRect(object: DrawingObject, viewport: Viewport, geom: GridGeometry): Rect {
  const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;
  const rect = anchorToRectPx(object.anchor, geom, zoom);
  const { scrollX, scrollY } = effectiveScrollForAnchor(object.anchor, viewport);
  const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
  const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
  return {
    x: rect.x - scrollX + headerOffsetX,
    y: rect.y - scrollY + headerOffsetY,
    width: rect.width,
    height: rect.height,
  };
}

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
  /** Precomputed rotation cos(theta) for hit testing (only meaningful when `transformFlags[i] !== 0`). */
  transformCos: Float64Array;
  /** Precomputed rotation sin(theta) for hit testing (only meaningful when `transformFlags[i] !== 0`). */
  transformSin: Float64Array;
  /**
   * Bit flags describing non-identity transforms.
   *
   * - bit0: flipH
   * - bit1: flipV
   * - bit2: has non-identity transform (rotation/flip)
   */
  transformFlags: Uint8Array;
  /**
   * Spatial bins keyed by bucket X then bucket Y.
   *
   * Each bucket stores indices into `ordered` / `bounds` in zOrder-desc order.
   */
  buckets: Map<number, Map<number, number[]>>;
  /** Indices of objects that were too large to bucket efficiently. */
  global: number[];
  bucketSizePx: number;
  /** Zoom used when computing `bounds`. */
  zoom: number;
  /** Geometry used to derive sheet-space bounds (also used for frozen-pane layout fallbacks). */
  geom: GridGeometry;
  /** Object id -> index into `ordered` / `bounds` (useful for selection/cursor logic). */
  byId: Map<number, number>;
}

export interface HitTestViewportLayout {
  frozenRows: number;
  frozenCols: number;
  headerOffsetX: number;
  headerOffsetY: number;
  frozenBoundaryX: number;
  frozenBoundaryY: number;
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
    /**
     * Zoom factor used when mapping EMU offsets/sizes to pixels.
     *
     * Defaults to 1, matching the overlay behavior when `viewport.zoom` is not provided.
     */
    zoom?: number;
  },
): HitTestIndex {
  const bucketSizePx = Math.max(1, Math.floor(opts?.bucketSizePx ?? 256));
  const maxBucketsPerObject = Math.max(1, Math.floor(opts?.maxBucketsPerObject ?? 256));
  const zoom = Number.isFinite(opts?.zoom) && (opts!.zoom as number) > 0 ? (opts!.zoom as number) : 1;
  // Walk from top to bottom (highest zOrder first).
  //
  // Perf: avoid allocating/sorting when callers already keep objects zOrder-sorted
  // ascending (back-to-front). In that common case we only need to reverse for
  // hit testing (top-to-bottom).
  //
  // We also handle the "already descending" case (common for some document
  // sources) without needing a full sort: if the list is zOrder-non-increasing
  // we only need to reverse within equal-zOrder runs to match render order.
  let isNonDecreasing = true;
  let isNonIncreasing = true;
  let hasEqual = false;
  for (let i = 1; i < objects.length; i += 1) {
    const prev = objects[i - 1]!.zOrder;
    const curr = objects[i]!.zOrder;
    if (prev > curr) isNonDecreasing = false;
    if (prev < curr) isNonIncreasing = false;
    if (prev === curr) hasEqual = true;
    if (!isNonDecreasing && !isNonIncreasing) break;
  }

  const ordered: DrawingObject[] = (() => {
    // Fast paths for common monotonic orderings (no sort).
    if (objects.length <= 1) return objects as DrawingObject[];
    if (isNonDecreasing) {
      const reversed = new Array<DrawingObject>(objects.length);
      for (let i = 0; i < objects.length; i += 1) {
        reversed[i] = objects[objects.length - 1 - i]!;
      }
      return reversed;
    }

    if (isNonIncreasing) {
      if (!hasEqual) {
        // Already top-to-bottom with unique zOrders.
        return objects as DrawingObject[];
      }

      // Reverse within equal-zOrder runs so ties are hit-tested in reverse render
      // order (render is stable within equal zOrder values).
      const out = new Array<DrawingObject>(objects.length);
      let write = 0;
      let start = 0;
      while (start < objects.length) {
        const z = objects[start]!.zOrder;
        let end = start + 1;
        while (end < objects.length && objects[end]!.zOrder === z) end += 1;
        for (let i = end - 1; i >= start; i -= 1) {
          out[write++] = objects[i]!;
        }
        start = end;
      }
      return out;
    }

    const sorted = [...objects].sort((a, b) => a.zOrder - b.zOrder);
    sorted.reverse();
    return sorted;
  })();
  const bounds: Rect[] = new Array(ordered.length);
  const aabbs: Rect[] = new Array(ordered.length);
  const transformCos = new Float64Array(ordered.length);
  const transformSin = new Float64Array(ordered.length);
  const transformFlags = new Uint8Array(ordered.length);
  const buckets: Map<number, Map<number, number[]>> = new Map();
  const global: number[] = [];
  const byId = new Map<number, number>();

  for (let i = 0; i < ordered.length; i += 1) {
    const obj = ordered[i]!;
    byId.set(obj.id, i);
    const rect = anchorToRectPx(obj.anchor, geom, zoom);
    bounds[i] = rect;
    let aabb = rect;
    let cos = 1;
    let sin = 0;
    let flags = 0;

    const transform = obj.transform;
    if (hasNonIdentityTransform(transform)) {
      flags = 4;
      if (transform!.flipH) flags |= 1;
      if (transform!.flipV) flags |= 2;
      const radians = (transform!.rotationDeg * Math.PI) / 180;
      cos = Math.cos(radians);
      sin = Math.sin(radians);
      aabb = rectToAabb(rect, cos, sin, flags);
    }

    transformCos[i] = cos;
    transformSin[i] = sin;
    transformFlags[i] = flags;
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

  return { ordered, bounds, aabbs, transformCos, transformSin, transformFlags, buckets, global, bucketSizePx, geom, byId, zoom };
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

    const flags = index.transformFlags[next] ?? 0;
    if ((flags & 4) !== 0) {
      const cos = index.transformCos[next] ?? 1;
      const sin = index.transformSin[next] ?? 0;
      if (!pointInTransformedRect(sheetX, sheetY, rect, cos, sin, flags)) continue;
    }

    return next;
  }

  return null;
}

export function hitTestDrawings(
  index: HitTestIndex,
  viewport: Viewport,
  x: number,
  y: number,
  geom: GridGeometry = index.geom,
  layout?: HitTestViewportLayout,
): HitTestResult | null {
  const headerOffsetX = layout
    ? layout.headerOffsetX
    : Number.isFinite(viewport.headerOffsetX)
      ? Math.max(0, viewport.headerOffsetX!)
      : 0;
  const headerOffsetY = layout
    ? layout.headerOffsetY
    : Number.isFinite(viewport.headerOffsetY)
      ? Math.max(0, viewport.headerOffsetY!)
      : 0;

  // Ignore pointer events over the header area; drawings are rendered under headers.
  if (x < headerOffsetX || y < headerOffsetY) return null;

  const frozenRows = layout
    ? layout.frozenRows
    : Number.isFinite(viewport.frozenRows)
      ? Math.max(0, Math.trunc(viewport.frozenRows!))
      : 0;
  const frozenCols = layout
    ? layout.frozenCols
    : Number.isFinite(viewport.frozenCols)
      ? Math.max(0, Math.trunc(viewport.frozenCols!))
      : 0;
  let frozenBoundaryX = layout ? layout.frozenBoundaryX : headerOffsetX;
  let frozenBoundaryY = layout ? layout.frozenBoundaryY : headerOffsetY;

  if (!layout && frozenCols > 0) {
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

  if (!layout && frozenRows > 0) {
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

  const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;
  if (Math.abs(zoom - index.zoom) > 1e-6) {
    // Zoom changes affect EMU->px conversions; fall back to a linear scan (still
    // respecting z-order) if the caller didn't rebuild the index for the current
    // zoom. This keeps hit testing consistent with `DrawingOverlay.render()`.
    const hasFrozenPanes = frozenRows !== 0 || frozenCols !== 0;
    for (const obj of index.ordered) {
      if (hasFrozenPanes) {
        const anchor = obj.anchor;
        const objInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < frozenRows;
        const objInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < frozenCols;
        if (objInFrozenRows !== inFrozenRows || objInFrozenCols !== inFrozenCols) continue;
      }

      const rect = anchorToRectPx(obj.anchor, geom, zoom);
      const transform = obj.transform;
      if (hasNonIdentityTransform(transform)) {
        let flags = 4;
        if (transform!.flipH) flags |= 1;
        if (transform!.flipV) flags |= 2;
        const radians = (transform!.rotationDeg * Math.PI) / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);
        const aabb = rectToAabb(rect, cos, sin, flags);
        if (!pointInRect(sheetX, sheetY, aabb)) continue;
        if (!pointInTransformedRect(sheetX, sheetY, rect, cos, sin, flags)) continue;
      } else if (!pointInRect(sheetX, sheetY, rect)) {
        continue;
      }

      const screen = {
        x: rect.x - scrollX + headerOffsetX,
        y: rect.y - scrollY + headerOffsetY,
        width: rect.width,
        height: rect.height,
      };
      return { object: obj, bounds: screen };
    }
    return null;
  }

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
  layout?: HitTestViewportLayout,
): DrawingObject | null {
  const headerOffsetX = layout
    ? layout.headerOffsetX
    : Number.isFinite(viewport.headerOffsetX)
      ? Math.max(0, viewport.headerOffsetX!)
      : 0;
  const headerOffsetY = layout
    ? layout.headerOffsetY
    : Number.isFinite(viewport.headerOffsetY)
      ? Math.max(0, viewport.headerOffsetY!)
      : 0;

  if (x < headerOffsetX || y < headerOffsetY) return null;

  const frozenRows = layout
    ? layout.frozenRows
    : Number.isFinite(viewport.frozenRows)
      ? Math.max(0, Math.trunc(viewport.frozenRows!))
      : 0;
  const frozenCols = layout
    ? layout.frozenCols
    : Number.isFinite(viewport.frozenCols)
      ? Math.max(0, Math.trunc(viewport.frozenCols!))
      : 0;

  let frozenBoundaryX = layout ? layout.frozenBoundaryX : headerOffsetX;
  let frozenBoundaryY = layout ? layout.frozenBoundaryY : headerOffsetY;

  if (!layout && frozenCols > 0) {
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

  if (!layout && frozenRows > 0) {
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

  const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;
  if (Math.abs(zoom - index.zoom) > 1e-6) {
    const hasFrozenPanes = frozenRows !== 0 || frozenCols !== 0;
    for (const obj of index.ordered) {
      if (hasFrozenPanes) {
        const anchor = obj.anchor;
        const objInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < frozenRows;
        const objInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < frozenCols;
        if (objInFrozenRows !== inFrozenRows || objInFrozenCols !== inFrozenCols) continue;
      }

      const rect = anchorToRectPx(obj.anchor, geom, zoom);
      const transform = obj.transform;
      if (hasNonIdentityTransform(transform)) {
        let flags = 4;
        if (transform!.flipH) flags |= 1;
        if (transform!.flipV) flags |= 2;
        const radians = (transform!.rotationDeg * Math.PI) / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);
        const aabb = rectToAabb(rect, cos, sin, flags);
        if (!pointInRect(sheetX, sheetY, aabb)) continue;
        if (!pointInTransformedRect(sheetX, sheetY, rect, cos, sin, flags)) continue;
      } else if (!pointInRect(sheetX, sheetY, rect)) {
        continue;
      }
      return obj;
    }
    return null;
  }

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

function pointInTransformedRect(x: number, y: number, rect: Rect, cos: number, sin: number, flags: number): boolean {
  if (!(rect.width > 0 && rect.height > 0)) return false;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const dx = x - cx;
  const dy = y - cy;
  // Inverse transform of: scale(flip) then rotate(theta).
  // Apply rotate(-theta) then scale(flip).
  let lx = dx * cos + dy * sin;
  let ly = -dx * sin + dy * cos;
  if (flags & 1) lx = -lx;
  if (flags & 2) ly = -ly;
  const hw = rect.width / 2;
  const hh = rect.height / 2;
  // Account for floating-point drift when `cos`/`sin` come from angles like 90°
  // where values are extremely close to 0/1. Without a tiny epsilon we can miss
  // hits that are conceptually on the boundary (e.g. resize-handle centers for
  // rotated objects).
  const eps = 1e-6;
  return lx >= -hw - eps && lx <= hw + eps && ly >= -hh - eps && ly <= hh + eps;
}

function rectToAabb(rect: Rect, cos: number, sin: number, flags: number): Rect {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;

  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  const visitCorner = (dx: number, dy: number) => {
    let x = dx;
    let y = dy;
    if (flags & 1) x = -x;
    if (flags & 2) y = -y;
    // Forward transform: scale(flip) then rotate(theta).
    const tx = x * cos - y * sin;
    const ty = x * sin + y * cos;
    const wx = cx + tx;
    const wy = cy + ty;
    if (wx < minX) minX = wx;
    if (wx > maxX) maxX = wx;
    if (wy < minY) minY = wy;
    if (wy > maxY) maxY = wy;
  };

  visitCorner(-hw, -hh);
  visitCorner(hw, -hh);
  visitCorner(hw, hh);
  visitCorner(-hw, hh);

  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}
