import type { DrawingObject, Rect } from "./types";
import { anchorToRectPx } from "./overlay";
import type { GridGeometry, Viewport } from "./overlay";

export interface HitTestResult {
  object: DrawingObject;
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
   * Sheet-space bounds (px) for each entry in `ordered`.
   *
   * These are computed once when the index is built so hit tests avoid repeatedly
   * calling `anchorToRectPx` on hot paths.
   */
  bounds: Rect[];
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
  const buckets: Map<number, Map<number, number[]>> = new Map();
  const global: number[] = [];
  const byId = new Map<number, number>();

  for (let i = 0; i < ordered.length; i += 1) {
    const obj = ordered[i]!;
    byId.set(obj.id, i);
    const rect = anchorToRectPx(obj.anchor, geom);
    bounds[i] = rect;

    const x1 = rect.x;
    const y1 = rect.y;
    const x2 = rect.x + rect.width;
    const y2 = rect.y + rect.height;

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

  return { ordered, bounds, buckets, global, bucketSizePx, geom, byId };
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

    const rect = index.bounds[next]!;
    if (hasFrozenPanes) {
      const anchor = index.ordered[next]!.anchor;
      const objInFrozenRows = anchor.type !== "absolute" && anchor.from.cell.row < frozenRows;
      const objInFrozenCols = anchor.type !== "absolute" && anchor.from.cell.col < frozenCols;
      // Excel-like pane routing: each drawing belongs to exactly one quadrant, so pointer
      // hits are constrained to the quadrant under the cursor.
      if (objInFrozenRows !== inFrozenRows || objInFrozenCols !== inFrozenCols) continue;
    }

    if (pointInRect(sheetX, sheetY, rect)) {
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

function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}
