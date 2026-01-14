import type { DrawingObject, ImageEntry, ImageStore } from "./types";

export type BitmapCacheLike = {
  invalidate(imageId: string): void;
};

export type ImagePersistenceLike = {
  delete(imageId: string): Promise<void>;
  set?(entry: ImageEntry): Promise<void>;
};

type GcCandidate = {
  imageId: string;
  gcAfterMs: number;
};

export type WorkbookImageManagerOptions = {
  images: ImageStore;
  /**
   * Optional decoded-bitmap cache (e.g. used by a DrawingOverlay) that should be invalidated when
   * we evict image bytes.
   */
  bitmapCache?: BitmapCacheLike;
  /**
   * Optional persistence store (IndexedDB) for image bytes.
   */
  persistence?: ImagePersistenceLike;
  /**
   * Minimum time to retain unreferenced image bytes, so a deletion can still be undone.
   */
  gcGracePeriodMs?: number;
};

const DEFAULT_GC_GRACE_PERIOD_MS = 5 * 60_000; // 5 minutes

/**
 * Maintains per-workbook image reference counts derived from per-sheet drawing metadata, and
 * garbage-collects image bytes when they are no longer referenced.
 *
 * To preserve undo/redo correctness, deletions are deferred for a grace period.
 */
export class WorkbookImageManager {
  readonly imageRefCount = new Map<string, number>();

  private readonly drawingsBySheetId = new Map<string, DrawingObject[]>();
  private readonly externalImageRefCount = new Map<string, number>();
  private readonly gcCandidates = new Map<string, GcCandidate>();
  private gcTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly gcGracePeriodMs: number;

  constructor(private readonly opts: WorkbookImageManagerOptions) {
    this.gcGracePeriodMs =
      typeof opts.gcGracePeriodMs === "number" && Number.isFinite(opts.gcGracePeriodMs) && opts.gcGracePeriodMs >= 0
        ? Math.trunc(opts.gcGracePeriodMs)
        : DEFAULT_GC_GRACE_PERIOD_MS;
  }

  /**
   * Replace all sheet drawings in one operation, then recompute refcounts once.
   */
  replaceAllSheetDrawings(
    drawingsBySheetId: Map<string, DrawingObject[]>,
    opts?: { externalImageIds?: Iterable<string> },
  ): void {
    this.drawingsBySheetId.clear();
    for (const [sheetId, objects] of drawingsBySheetId.entries()) {
      const id = String(sheetId ?? "");
      if (!id) continue;
      this.drawingsBySheetId.set(id, Array.isArray(objects) ? objects : []);
    }

    if (opts?.externalImageIds) {
      this.externalImageRefCount.clear();
      for (const raw of opts.externalImageIds) {
        const id = String(raw ?? "");
        if (!id) continue;
        this.externalImageRefCount.set(id, (this.externalImageRefCount.get(id) ?? 0) + 1);
      }
    }

    this.recomputeImageRefCounts();
  }

  /**
   * Set additional non-drawing references (e.g. sheet background images) that should prevent image GC.
   */
  setExternalImageIds(ids: Iterable<string>): void {
    this.externalImageRefCount.clear();
    for (const raw of ids) {
      const id = String(raw ?? "");
      if (!id) continue;
      this.externalImageRefCount.set(id, (this.externalImageRefCount.get(id) ?? 0) + 1);
    }
    this.recomputeImageRefCounts();
  }

  setSheetDrawings(sheetId: string, objects: DrawingObject[]): void {
    const id = String(sheetId ?? "");
    if (!id) return;
    this.drawingsBySheetId.set(id, Array.isArray(objects) ? objects : []);
    this.recomputeImageRefCounts();
  }

  deleteSheet(sheetId: string): void {
    const id = String(sheetId ?? "");
    if (!id) return;
    if (!this.drawingsBySheetId.has(id)) return;
    this.drawingsBySheetId.delete(id);
    this.recomputeImageRefCounts();
  }

  /**
   * Recompute image ref counts from the current sheet drawings state.
   *
   * This is currently O(total drawings) but drawing counts are expected to be small compared to
   * cell grid operations, and simplicity helps correctness.
   */
  recomputeImageRefCounts(): void {
    const next = new Map<string, number>();

    for (const [imageId, count] of this.externalImageRefCount.entries()) {
      if (count > 0) next.set(imageId, (next.get(imageId) ?? 0) + count);
    }

    for (const objects of this.drawingsBySheetId.values()) {
      if (!Array.isArray(objects)) continue;
      for (const obj of objects) {
        if (!obj || obj.kind?.type !== "image") continue;
        const imageId = String((obj.kind as any).imageId ?? "");
        if (!imageId) continue;
        next.set(imageId, (next.get(imageId) ?? 0) + 1);
      }
    }

    const prev = new Map(this.imageRefCount);
    this.imageRefCount.clear();
    for (const [imageId, count] of next.entries()) {
      if (count > 0) this.imageRefCount.set(imageId, count);
    }

    // Queue GC for images that lost their last reference.
    for (const [imageId, prevCount] of prev.entries()) {
      const nextCount = next.get(imageId) ?? 0;
      if (prevCount > 0 && nextCount === 0) {
        this.queueGc(imageId);
      }
    }

    // Cancel GC for images that regained a reference.
    for (const [imageId, nextCount] of next.entries()) {
      if (nextCount <= 0) continue;
      if ((prev.get(imageId) ?? 0) === 0) {
        this.gcCandidates.delete(imageId);
      }
    }

    this.rescheduleGcTimer();
  }

  private queueGc(imageId: string): void {
    const id = String(imageId ?? "");
    if (!id) return;

    const now = Date.now();
    this.gcCandidates.set(id, { imageId: id, gcAfterMs: now + this.gcGracePeriodMs });
  }

  private rescheduleGcTimer(): void {
    if (this.gcTimer) {
      clearTimeout(this.gcTimer);
      this.gcTimer = null;
    }

    if (this.gcCandidates.size === 0) return;

    const now = Date.now();
    let nextAt = Number.POSITIVE_INFINITY;
    for (const c of this.gcCandidates.values()) {
      nextAt = Math.min(nextAt, c.gcAfterMs);
    }
    if (!Number.isFinite(nextAt)) return;

    const delay = Math.max(0, nextAt - now);
    this.gcTimer = setTimeout(() => {
      this.gcTimer = null;
      void this.runGcNow();
    }, delay);
  }

  async runGcNow(opts: { force?: boolean } = {}): Promise<void> {
    const force = Boolean(opts.force);
    const now = Date.now();
    const toDelete: string[] = [];

    for (const candidate of this.gcCandidates.values()) {
      if (!force && candidate.gcAfterMs > now) continue;
      // Only GC if still unreferenced.
      if ((this.imageRefCount.get(candidate.imageId) ?? 0) > 0) continue;
      toDelete.push(candidate.imageId);
    }

    if (toDelete.length === 0) {
      this.rescheduleGcTimer();
      return;
    }

    for (const imageId of toDelete) {
      // Delete from in-memory store.
      this.opts.images.delete(imageId);
      // Drop decoded bitmap cache entry (if present).
      this.opts.bitmapCache?.invalidate(imageId);
      // Delete from persistence.
      if (this.opts.persistence) {
        try {
          await this.opts.persistence.delete(imageId);
        } catch {
          // Best-effort. The in-memory store is the primary source for rendering.
        }
      }
      this.gcCandidates.delete(imageId);
    }

    this.rescheduleGcTimer();
  }

  /**
   * Store image bytes in the in-memory store and (optionally) persist them.
   */
  async putImage(entry: ImageEntry): Promise<void> {
    this.opts.images.set(entry);
    if (this.opts.persistence?.set) {
      await this.opts.persistence.set(entry);
    }
  }

  dispose(): void {
    if (this.gcTimer) {
      clearTimeout(this.gcTimer);
      this.gcTimer = null;
    }
    this.gcCandidates.clear();
    this.drawingsBySheetId.clear();
    this.externalImageRefCount.clear();
    this.imageRefCount.clear();
  }
}
