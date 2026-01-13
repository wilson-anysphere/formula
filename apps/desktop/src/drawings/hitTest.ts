import type { DrawingObject, Rect } from "./types";
import { anchorToRectPx } from "./overlay";
import type { GridGeometry, Viewport } from "./overlay";

export interface HitTestResult {
  object: DrawingObject;
  bounds: Rect;
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

  for (let i = 0; i < ordered.length; i += 1) {
    const obj = ordered[i]!;
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

  return { ordered, bounds, buckets, global, bucketSizePx };
}

export function hitTestDrawings(index: HitTestIndex, viewport: Viewport, x: number, y: number): HitTestResult | null {
  // Convert from screen-space to sheet-space.
  const sheetX = x + viewport.scrollX;
  const sheetY = y + viewport.scrollY;

  const bx = Math.floor(sheetX / index.bucketSizePx);
  const by = Math.floor(sheetY / index.bucketSizePx);

  const bucket = index.buckets.get(bx)?.get(by);

  // Merge the bucket-specific list and global list (both sorted in zOrder-desc order because we
  // inserted indices in that order).
  const bucketList = bucket ?? EMPTY_LIST;
  const globalList = index.global.length > 0 ? index.global : EMPTY_LIST;

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
    if (pointInRect(sheetX, sheetY, rect)) {
      const screen = {
        x: rect.x - viewport.scrollX,
        y: rect.y - viewport.scrollY,
        width: rect.width,
        height: rect.height,
      };
      return { object: index.ordered[next]!, bounds: screen };
    }
  }

  return null;
}

const EMPTY_LIST: number[] = [];

function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}
