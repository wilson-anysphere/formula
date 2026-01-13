import type { ImageEntry } from "./types";

/**
 * Cache for decoded images.
 *
 * `createImageBitmap` is asynchronous and relatively expensive, so we keep a
 * Promise per imageId to dedupe concurrent requests.
 */
export class ImageBitmapCache {
  /**
   * LRU entries keyed by image id.
   *
   * `Map` insertion order represents recency:
   * - newest entries are at the end
   * - oldest entries are at the start
   *
   * Entries exist as soon as `get()` is called, even while the decode is still
   * pending.
   *
   * Note: we also "touch" (move-to-end) entries when their decode finishes so a
   * late-resolving request isn't immediately evicted+closed before its callers
   * observe the resolved bitmap (our internal `.then` handlers run before any
   * caller `await`/`.then` handlers attached later).
   */
  private readonly entries = new Map<
    string,
    {
      promise: Promise<ImageBitmap>;
      bitmap?: ImageBitmap;
    }
  >();
  private maxEntries: number;
  private decodedCount = 0;

  constructor(options?: { maxEntries?: number }) {
    const max = options?.maxEntries;
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(max ?? 256);
  }

  setMaxEntries(maxEntries: number): void {
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(maxEntries);
    this.evictIfNeeded();
  }

  get(entry: ImageEntry): Promise<ImageBitmap> {
    const id = entry.id;
    const existing = this.entries.get(id);
    if (existing) {
      // Mark as most-recently-used.
      this.entries.delete(id);
      this.entries.set(id, existing);
      return existing.promise;
    }

    const promise = ImageBitmapCache.decode(entry);
    const record = { promise, bitmap: undefined as ImageBitmap | undefined };
    this.entries.set(id, record);

    // Once the decode finishes, populate the LRU cache *only if* this promise is
    // still the active in-flight request for the image id. This prevents stale
    // requests (e.g. after `invalidate()` or `clear()`) from repopulating the
    // cache.
    void promise.then(
      (bitmap) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) return;

        // A max size of 0 means caching is disabled. Still dedupe the in-flight
        // request, but don't store the decoded bitmap (and do not close it,
        // since the caller owns the resolved value).
        if (this.maxEntries === 0) {
          this.entries.delete(id);
          return;
        }

        // This is the first time the entry has resolved; attach the bitmap and
        // count it toward the cache size.
        if (!record.bitmap) {
          record.bitmap = bitmap;
          this.decodedCount++;
        } else if (record.bitmap !== bitmap) {
          // Should be impossible, but keep accounting correct if it happens.
          ImageBitmapCache.tryClose(record.bitmap);
          record.bitmap = bitmap;
        }

        // Mark as most-recently-used. This avoids evicting+closing the bitmap in
        // the same microtask that resolves the promise (which would make the
        // resolved bitmap unusable for awaiting callers).
        this.entries.delete(id);
        this.entries.set(id, record);

        this.evictIfNeeded();
      },
      () => {
        const current = this.entries.get(id);
        if (current === record && current.promise === promise) {
          this.entries.delete(id);
        }
      },
    );

    return promise;
  }

  invalidate(imageId: string): void {
    const existing = this.entries.get(imageId);
    if (!existing) return;

    this.entries.delete(imageId);
    if (existing.bitmap) {
      this.decodedCount--;
      ImageBitmapCache.tryClose(existing.bitmap);
    }
  }

  clear(): void {
    for (const entry of this.entries.values()) {
      if (entry.bitmap) ImageBitmapCache.tryClose(entry.bitmap);
    }
    this.entries.clear();
    this.decodedCount = 0;
  }

  private evictIfNeeded(): void {
    while (this.decodedCount > this.maxEntries) {
      let evicted = false;
      for (const [id, entry] of this.entries) {
        if (!entry.bitmap) continue;
        this.entries.delete(id);
        this.decodedCount--;
        ImageBitmapCache.tryClose(entry.bitmap);
        evicted = true;
        break;
      }
      if (!evicted) return;
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
