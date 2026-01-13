import type { ImageEntry } from "./types";

/**
 * Cache for decoded images.
 *
 * `createImageBitmap` is asynchronous and relatively expensive, so we keep a
 * Promise per imageId to dedupe concurrent requests.
 */
export class ImageBitmapCache {
  /**
   * Decoded bitmap LRU. Map insertion order is used for recency tracking:
   * - newest entries are at the end
   * - oldest entries are at the start
   */
  private readonly bitmaps = new Map<string, ImageBitmap>();
  /** In-flight decodes, used to dedupe concurrent `get()` calls. */
  private readonly inflight = new Map<string, Promise<ImageBitmap>>();
  private maxEntries: number;

  constructor(options?: { maxEntries?: number }) {
    const max = options?.maxEntries;
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(max ?? 256);
  }

  setMaxEntries(maxEntries: number): void {
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(maxEntries);
    this.evictIfNeeded();
  }

  get(entry: ImageEntry): Promise<ImageBitmap> {
    const cached = this.bitmaps.get(entry.id);
    if (cached) {
      // Mark as most-recently-used.
      this.bitmaps.delete(entry.id);
      this.bitmaps.set(entry.id, cached);
      return Promise.resolve(cached);
    }

    const existing = this.inflight.get(entry.id);
    if (existing) return existing;

    const id = entry.id;
    const promise = ImageBitmapCache.decode(entry);
    this.inflight.set(id, promise);

    // Once the decode finishes, populate the LRU cache *only if* this promise is
    // still the active in-flight request for the image id. This prevents stale
    // requests (e.g. after `invalidate()` or `clear()`) from repopulating the
    // cache.
    void promise.then(
      (bitmap) => {
        if (this.inflight.get(id) !== promise) return;
        this.inflight.delete(id);

        // A max size of 0 means caching is disabled. Still dedupe the in-flight
        // request, but don't store the decoded bitmap (and do not close it,
        // since the caller owns the resolved value).
        if (this.maxEntries === 0) return;

        // Insert as most-recently-used. If something already existed (should be
        // rare, but can happen due to races), close it before replacing.
        const prior = this.bitmaps.get(id);
        if (prior) ImageBitmapCache.tryClose(prior);
        this.bitmaps.delete(id);
        this.bitmaps.set(id, bitmap);
        this.evictIfNeeded();
      },
      () => {
        if (this.inflight.get(id) === promise) {
          this.inflight.delete(id);
        }
      },
    );

    return promise;
  }

  invalidate(imageId: string): void {
    // Removing an entry should drop the cached bitmap (if any) and ensure a
    // pending decode can't later repopulate the cache.
    const existing = this.bitmaps.get(imageId);
    if (existing) {
      this.bitmaps.delete(imageId);
      ImageBitmapCache.tryClose(existing);
    }
    this.inflight.delete(imageId);
  }

  clear(): void {
    for (const bitmap of this.bitmaps.values()) {
      ImageBitmapCache.tryClose(bitmap);
    }
    this.bitmaps.clear();
    this.inflight.clear();
  }

  private evictIfNeeded(): void {
    while (this.bitmaps.size > this.maxEntries) {
      const oldest = this.bitmaps.entries().next().value as [string, ImageBitmap] | undefined;
      if (!oldest) return;
      const [id, bitmap] = oldest;
      this.bitmaps.delete(id);
      ImageBitmapCache.tryClose(bitmap);
    }
  }

  private static normalizeMaxEntries(value: number): number {
    if (!Number.isFinite(value)) return 256;
    // Clamp to a small-ish non-negative integer to avoid pathological behavior.
    return Math.max(0, Math.floor(value));
  }

  private static decode(entry: ImageEntry): Promise<ImageBitmap> {
    const buffer = new ArrayBuffer(entry.bytes.byteLength);
    new Uint8Array(buffer).set(entry.bytes);
    const blob = new Blob([buffer], { type: entry.mimeType });
    return createImageBitmap(blob);
  }

  private static tryClose(bitmap: ImageBitmap): void {
    try {
      const anyBitmap = bitmap as any;
      if (anyBitmap && typeof anyBitmap.close === "function") {
        anyBitmap.close();
      }
    } catch {
      // Ignore close errors (some runtimes may throw or not implement it).
    }
  }
}
