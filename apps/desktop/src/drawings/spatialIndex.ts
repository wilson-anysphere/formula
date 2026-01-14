import type { Anchor, DrawingObject, DrawingTransform, Rect } from "./types";
import type { GridGeometry } from "./overlay";

import { emuToPx } from "../shared/emu.js";

export const DEFAULT_DRAWING_SPATIAL_INDEX_TILE_SIZE_PX = 512;

const A1_CELL = { row: 0, col: 0 };

function anchorToRectPx(anchor: Anchor, geom: GridGeometry, zoom = 1): Rect {
  const scale = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  switch (anchor.type) {
    case "oneCell": {
      const origin = geom.cellOriginPx(anchor.from.cell);
      return {
        x: origin.x + emuToPx(anchor.from.offset.xEmu) * scale,
        y: origin.y + emuToPx(anchor.from.offset.yEmu) * scale,
        width: emuToPx(anchor.size.cx) * scale,
        height: emuToPx(anchor.size.cy) * scale,
      };
    }
    case "twoCell": {
      const fromOrigin = geom.cellOriginPx(anchor.from.cell);
      const toOrigin = geom.cellOriginPx(anchor.to.cell);

      const x1 = fromOrigin.x + emuToPx(anchor.from.offset.xEmu) * scale;
      const y1 = fromOrigin.y + emuToPx(anchor.from.offset.yEmu) * scale;
      // In DrawingML, `to` specifies the cell containing the bottom-right corner
      // (i.e. the first cell outside the object when the corner lands on a grid
      // boundary). The absolute end point is therefore the origin of the `to`
      // cell plus the offsets.
      const x2 = toOrigin.x + emuToPx(anchor.to.offset.xEmu) * scale;
      const y2 = toOrigin.y + emuToPx(anchor.to.offset.yEmu) * scale;

      return {
        x: Math.min(x1, x2),
        y: Math.min(y1, y2),
        width: Math.abs(x2 - x1),
        height: Math.abs(y2 - y1),
      };
    }
    case "absolute":
      // In DrawingML, absolute anchors are worksheet-space coordinates whose
      // origin is the top-left of the cell grid (A1), not the top-left of the
      // overall grid UI root (which may include row/column headers).
      //
      // Use the A1 origin from the provided grid geometry so absolute anchors
      // align with oneCell/twoCell anchors in shared-grid mode.
      const origin = geom.cellOriginPx(A1_CELL);
      return {
        x: origin.x + emuToPx(anchor.pos.xEmu) * scale,
        y: origin.y + emuToPx(anchor.pos.yEmu) * scale,
        width: emuToPx(anchor.size.cx) * scale,
        height: emuToPx(anchor.size.cy) * scale,
      };
  }
}

function intersects(a: Rect, b: Rect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function pointInRect(x: number, y: number, rect: Rect): boolean {
  return x >= rect.x && y >= rect.y && x <= rect.x + rect.width && y <= rect.y + rect.height;
}

function hasNonIdentityTransform(transform: DrawingTransform | undefined): boolean {
  if (!transform) return false;
  return transform.rotationDeg !== 0 || transform.flipH || transform.flipV;
}

function rectToAabb(rect: Rect, transform: DrawingTransform): Rect {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const hw = rect.width / 2;
  const hh = rect.height / 2;

  const radians = (transform.rotationDeg * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  const visitCorner = (dx: number, dy: number) => {
    let x = dx;
    let y = dy;
    if (transform.flipH) x = -x;
    if (transform.flipV) y = -y;
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

function pointInTransformedRect(x: number, y: number, rect: Rect, transform: DrawingTransform): boolean {
  if (!(rect.width > 0 && rect.height > 0)) return false;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const dx = x - cx;
  const dy = y - cy;
  // Inverse transform of: scale(flip) then rotate(theta).
  // Apply rotate(-theta) then scale(flip).
  const radians = (transform.rotationDeg * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  let lx = dx * cos + dy * sin;
  let ly = -dx * sin + dy * cos;
  if (transform.flipH) lx = -lx;
  if (transform.flipV) ly = -ly;
  const hw = rect.width / 2;
  const hh = rect.height / 2;
  return lx >= -hw && lx <= hw && ly >= -hh && ly <= hh;
}

export interface DrawingSpatialIndexHitTestResult {
  object: DrawingObject;
  /**
   * Screen-space bounds (CSS px) for the hit object (suitable for selection handles).
   */
  bounds: Rect;
}

/**
 * Uniform-grid spatial index for `DrawingObject`s.
 *
 * - Built in sheet-space pixels (same coordinate system as `GridGeometry`)
 * - Buckets are keyed by coarse tile coords (tile size defaults to ~512px)
 * - Per-bucket id lists are stored in z-order so queries don't need to sort.
 */
export class DrawingSpatialIndex {
  private readonly tileSize: number;
  private readonly maxBucketsPerObject: number;

  // buckets[x][y] => ordered ids
  private readonly buckets = new Map<number, Map<number, number[]>>();
  private readonly globalBucket: number[] = [];
  private readonly rectById = new Map<number, Rect>();
  private readonly aabbById = new Map<number, Rect>();
  private readonly objectById = new Map<number, DrawingObject>();
  private readonly orderById = new Map<number, number>();

  private lastObjects: DrawingObject[] | null = null;
  private lastGeom: GridGeometry | null = null;
  private lastZoom = 1;
  private dirty = true;

  // Query scratch (reused to avoid per-call allocations).
  private readonly bucketArraysScratch: number[][] = [];
  private readonly pointersScratch: number[] = [];
  private readonly seenGenerationById = new Map<number, number>();
  private seenGeneration = 1;

  constructor(opts?: { tileSizePx?: number; maxBucketsPerObject?: number }) {
    const tileSizePx = opts?.tileSizePx ?? DEFAULT_DRAWING_SPATIAL_INDEX_TILE_SIZE_PX;
    if (!Number.isFinite(tileSizePx) || tileSizePx <= 0) {
      throw new Error(`Invalid DrawingSpatialIndex tileSizePx: ${String(tileSizePx)}`);
    }
    this.tileSize = tileSizePx;
    const maxBucketsPerObject = opts?.maxBucketsPerObject ?? 256;
    this.maxBucketsPerObject = Math.max(1, Math.floor(maxBucketsPerObject));
  }

  /**
   * Marks the index dirty so the next `rebuild()` recomputes bucket membership.
   *
   * Useful when `GridGeometry` is stable by reference but cell sizes/origins
   * change (e.g. axis resize, zoom changes applied internally).
   */
  invalidate(): void {
    this.dirty = true;
  }

  /**
   * Rebuilds the entire index (unless inputs are unchanged and the index is not dirty).
   */
  rebuild(objects: DrawingObject[], geom: GridGeometry, zoom = 1): void {
    if (!this.dirty && this.lastObjects === objects && this.lastGeom === geom && this.lastZoom === zoom) {
      return;
    }

    this.lastObjects = objects;
    this.lastGeom = geom;
    this.lastZoom = zoom;
    this.dirty = false;

    this.buckets.clear();
    this.globalBucket.length = 0;
    this.rectById.clear();
    this.aabbById.clear();
    this.objectById.clear();
    this.orderById.clear();
    this.seenGenerationById.clear();
    this.seenGeneration = 1;

    // Sort once at rebuild time so per-frame queries don't sort.
    let sorted = true;
    for (let i = 1; i < objects.length; i += 1) {
      if (objects[i - 1]!.zOrder > objects[i]!.zOrder) {
        sorted = false;
        break;
      }
    }
    const ordered = sorted ? objects : [...objects].sort((a, b) => a.zOrder - b.zOrder);

    const tileSize = this.tileSize;
    const maxBucketsPerObject = this.maxBucketsPerObject;
    for (let i = 0; i < ordered.length; i += 1) {
      const obj = ordered[i]!;
      const rect = anchorToRectPx(obj.anchor, geom, zoom);
      const aabb = hasNonIdentityTransform(obj.transform) ? rectToAabb(rect, obj.transform!) : rect;

      this.rectById.set(obj.id, rect);
      this.aabbById.set(obj.id, aabb);
      this.objectById.set(obj.id, obj);
      this.orderById.set(obj.id, i);

      const x1 = aabb.x;
      const y1 = aabb.y;
      const x2 = aabb.x + aabb.width;
      const y2 = aabb.y + aabb.height;

      const minTileX = Math.floor(x1 / tileSize);
      const maxTileX = Math.floor(x2 / tileSize);
      const minTileY = Math.floor(y1 / tileSize);
      const maxTileY = Math.floor(y2 / tileSize);

      const bucketsWide = maxTileX - minTileX + 1;
      const bucketsHigh = maxTileY - minTileY + 1;
      const bucketCount = bucketsWide * bucketsHigh;

      if (!Number.isFinite(bucketCount) || bucketCount > maxBucketsPerObject) {
        this.globalBucket.push(obj.id);
        continue;
      }

      for (let tx = minTileX; tx <= maxTileX; tx += 1) {
        let col = this.buckets.get(tx);
        if (!col) {
          col = new Map<number, number[]>();
          this.buckets.set(tx, col);
        }
        for (let ty = minTileY; ty <= maxTileY; ty += 1) {
          let bucket = col.get(ty);
          if (!bucket) {
            bucket = [];
            col.set(ty, bucket);
          }
          bucket.push(obj.id);
        }
      }
    }
  }

  getObject(id: number): DrawingObject | null {
    return this.objectById.get(id) ?? null;
  }

  getRect(id: number): Rect | null {
    return this.rectById.get(id) ?? null;
  }

  getAabb(id: number): Rect | null {
    return this.aabbById.get(id) ?? null;
  }

  /**
   * Returns candidate objects whose bounding rect intersects the viewport.
   *
   * Returned in z-order (ascending).
   */
  query(viewportRectSheetSpace: Rect): DrawingObject[] {
    const tileSize = this.tileSize;
    const minTileX = Math.floor(viewportRectSheetSpace.x / tileSize);
    const maxTileX = Math.floor((viewportRectSheetSpace.x + viewportRectSheetSpace.width) / tileSize);
    const minTileY = Math.floor(viewportRectSheetSpace.y / tileSize);
    const maxTileY = Math.floor((viewportRectSheetSpace.y + viewportRectSheetSpace.height) / tileSize);

    const bucketArrays = this.bucketArraysScratch;
    bucketArrays.length = 0;

    for (let tx = minTileX; tx <= maxTileX; tx += 1) {
      const col = this.buckets.get(tx);
      if (!col) continue;
      for (let ty = minTileY; ty <= maxTileY; ty += 1) {
        const bucket = col.get(ty);
        if (bucket && bucket.length > 0) bucketArrays.push(bucket);
      }
    }

    if (this.globalBucket.length > 0) bucketArrays.push(this.globalBucket);

    if (bucketArrays.length === 0) return [];

    // Ensure `pointersScratch` has enough entries and is zeroed.
    const pointers = this.pointersScratch;
    if (pointers.length < bucketArrays.length) {
      pointers.length = bucketArrays.length;
    }
    for (let i = 0; i < bucketArrays.length; i += 1) pointers[i] = 0;

    // Bump generation for dedupe; reset if we overflow (extremely unlikely).
    this.seenGeneration += 1;
    if (this.seenGeneration === 0x7fffffff) {
      this.seenGeneration = 1;
      this.seenGenerationById.clear();
    }
    const generation = this.seenGeneration;

    const seen = this.seenGenerationById;
    const orderById = this.orderById;
    const aabbById = this.aabbById;
    const objectById = this.objectById;

    const out: DrawingObject[] = [];

    // K-way merge of per-bucket z-ordered id lists.
    while (true) {
      let bestBucketIndex = -1;
      let bestId = 0;
      let bestOrder = 0;

      for (let i = 0; i < bucketArrays.length; i += 1) {
        const bucket = bucketArrays[i]!;
        const p = pointers[i]!;
        if (p >= bucket.length) continue;
        const id = bucket[p]!;
        const order = orderById.get(id);
        if (order == null) continue;
        if (bestBucketIndex === -1 || order < bestOrder) {
          bestBucketIndex = i;
          bestId = id;
          bestOrder = order;
        }
      }

      if (bestBucketIndex === -1) break;
      pointers[bestBucketIndex]! += 1;

      if (seen.get(bestId) === generation) continue;
      seen.set(bestId, generation);

      const aabb = aabbById.get(bestId);
      if (!aabb) continue;
      if (!intersects(aabb, viewportRectSheetSpace)) continue;
      const obj = objectById.get(bestId);
      if (obj) out.push(obj);
    }

    return out;
  }

  /**
   * Performs a point hit-test (screen space) against objects in the relevant tile.
   *
   * Returns the topmost matching object (highest z-order).
   */
  hitTest(pointScreen: { x: number; y: number }, viewport: { scrollX: number; scrollY: number }): DrawingObject | null {
    const sheetX = viewport.scrollX + pointScreen.x;
    const sheetY = viewport.scrollY + pointScreen.y;
    // Use a point query on AABBs to collect a small candidate set, then test
    // the actual (potentially rotated/flipped) rectangle.
    const candidates = this.query({ x: sheetX, y: sheetY, width: 0, height: 0 });
    for (let i = candidates.length - 1; i >= 0; i -= 1) {
      const obj = candidates[i]!;
      const rect = this.rectById.get(obj.id);
      if (!rect) continue;
      if (hasNonIdentityTransform(obj.transform)) {
        if (!pointInTransformedRect(sheetX, sheetY, rect, obj.transform!)) continue;
      } else {
        if (!pointInRect(sheetX, sheetY, rect)) continue;
      }
      return obj;
    }
    return null;
  }

  hitTestWithBounds(
    pointScreen: { x: number; y: number },
    viewport: { scrollX: number; scrollY: number },
  ): DrawingSpatialIndexHitTestResult | null {
    const obj = this.hitTest(pointScreen, viewport);
    if (!obj) return null;
    const rect = this.rectById.get(obj.id);
    if (!rect) return null;
    return {
      object: obj,
      bounds: {
        x: rect.x - viewport.scrollX,
        y: rect.y - viewport.scrollY,
        width: rect.width,
        height: rect.height,
      },
    };
  }
}
