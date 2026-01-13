import type { ImageEntry } from "./types";

export interface ImageBitmapCacheOptions {
  /**
   * Maximum number of decoded bitmaps to keep in the in-memory LRU cache.
   *
   * A value of 0 disables bitmap caching (but still dedupes concurrent decodes).
   *
   * @defaultValue 256
   */
  maxEntries?: number;

  /**
   * How long (in ms) to keep a short-lived "negative cache" entry after a decode
   * failure.
   *
   * This helps avoid tight retry loops (e.g. every render pass) when an image is
   * corrupted or uses an unsupported mime type, while still allowing retries
   * shortly after.
   *
   * Set to 0 to disable.
   *
   * @defaultValue 0
   */
  negativeCacheMs?: number;
}

export interface ImageBitmapCacheGetOptions {
  /**
   * Abort signal for callers that no longer need the bitmap (e.g. offscreen
   * images). Aborting rejects the returned promise with an `AbortError` and
   * ensures the cache entry is cleaned up so future callers can retry.
   */
  signal?: AbortSignal;
}

type CacheEntry = {
  promise: Promise<ImageBitmap>;
  bitmap?: ImageBitmap;
  /**
   * Number of signal-based waiters currently attached to this in-flight decode.
   *
   * This lets us avoid dropping the cache entry when *one* caller aborts but
   * another caller is still awaiting the same decode (e.g. overlapping render
   * passes).
   */
  waiters: number;
  /**
   * True if any consumer requested the bitmap without an AbortSignal while the
   * decode is still in-flight.
   *
   * These consumers are not tracked as `waiters`, so we conservatively treat the
   * decode as wanted and keep the cache entry.
   */
  pinned: boolean;
};

type NegativeCacheEntry = {
  error: unknown;
  expiresAt: number;
};

function createAbortError(): Error {
  const err = new Error("The operation was aborted.");
  err.name = "AbortError";
  return err;
}

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
  private readonly entries = new Map<string, CacheEntry>();
  private maxEntries: number;
  private decodedCount = 0;

  private readonly negativeCache = new Map<string, NegativeCacheEntry>();
  private readonly negativeCacheMs: number;

  /**
   * Test/debug counter for how many decode attempts have failed.
   *
   * This is intentionally not private so unit tests can assert that failures are
   * observed and handled.
   */
  __testOnly_failCount = 0;

  constructor(options: ImageBitmapCacheOptions = {}) {
    const max = options.maxEntries;
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(max ?? 256);
    this.negativeCacheMs = options.negativeCacheMs ?? 0;
  }

  setMaxEntries(maxEntries: number): void {
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(maxEntries);
    this.evictIfNeeded();
  }

  get(entry: ImageEntry, opts: ImageBitmapCacheGetOptions = {}): Promise<ImageBitmap> {
    const id = entry.id;
    const existing = this.entries.get(id);
    if (existing) {
      // Mark as most-recently-used.
      this.entries.delete(id);
      this.entries.set(id, existing);
      if (!existing.bitmap) {
        if (opts.signal) existing.waiters++;
        else existing.pinned = true;
      }
      return this.wrapWithAbort(id, existing, opts.signal, Boolean(opts.signal) && !existing.bitmap);
    }

    // If the request is already aborted, don't start decoding (and don't poison
    // the cache with a rejected promise).
    if (opts.signal?.aborted) {
      return Promise.reject(createAbortError());
    }

    const cachedFailure = this.negativeCache.get(id);
    if (cachedFailure) {
      if (cachedFailure.expiresAt > Date.now()) {
        return Promise.reject(cachedFailure.error);
      }
      this.negativeCache.delete(id);
    }

    const promise = ImageBitmapCache.decode(entry);
    const record: CacheEntry = {
      promise,
      bitmap: undefined as ImageBitmap | undefined,
      waiters: opts.signal ? 1 : 0,
      pinned: !opts.signal,
    };
    this.entries.set(id, record);

    // Once the decode finishes, populate the LRU cache *only if* this promise is
    // still the active in-flight request for the image id. This prevents stale
    // requests (e.g. after `invalidate()` or `clear()`, or abort cleanup) from
    // repopulating the cache.
    void promise.then(
      (bitmap) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) return;

        // If the decode finishes after all callers have aborted (and no
        // untracked waiters exist), drop it immediately to avoid caching a bitmap
        // nobody will use.
        if (!record.pinned && record.waiters === 0) {
          this.entries.delete(id);
          ImageBitmapCache.tryClose(bitmap);
          return;
        }

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

        this.negativeCache.delete(id);

        // Mark as most-recently-used. This avoids evicting+closing the bitmap in
        // the same microtask that resolves the promise (which would make the
        // resolved bitmap unusable for awaiting callers).
        this.entries.delete(id);
        this.entries.set(id, record);

        this.evictIfNeeded();
      },
      (err) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) return;

        this.__testOnly_failCount++;
        if (this.negativeCacheMs > 0) {
          this.negativeCache.set(id, { error: err, expiresAt: Date.now() + this.negativeCacheMs });
        }

        this.entries.delete(id);
      },
    );

    return this.wrapWithAbort(id, record, opts.signal, Boolean(opts.signal));
  }

  invalidate(imageId: string): void {
    const existing = this.entries.get(imageId);
    if (!existing) {
      this.negativeCache.delete(imageId);
      return;
    }

    this.entries.delete(imageId);
    this.negativeCache.delete(imageId);
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
    this.negativeCache.clear();
  }

  private wrapWithAbort(
    imageId: string,
    record: CacheEntry,
    signal: AbortSignal | undefined,
    trackWaiter: boolean,
  ): Promise<ImageBitmap> {
    if (!signal) return record.promise;

    let released = false;
    const release = () => {
      if (!trackWaiter) return;
      if (released) return;
      released = true;
      record.waiters = Math.max(0, record.waiters - 1);
    };

    const abortAndCleanup = (): Error => {
      release();
      // Do not immediately delete the cache entry here.
      //
      // Multiple callers can share an in-flight decode promise; if one aborts we
      // still want others (or a subsequent render pass) to be able to reuse the
      // decode. If everyone aborts, the `.then` handler attached in `get()`
      // cleans up the entry once the decode finishes.
      return createAbortError();
    };

    if (signal.aborted) {
      return Promise.reject(abortAndCleanup());
    }

    // Note: `createImageBitmap` itself is not abortable; we only provide
    // predictable caller semantics and cache cleanup.
    return new Promise<ImageBitmap>((resolve, reject) => {
      const onAbort = () => {
        reject(abortAndCleanup());
      };

      signal.addEventListener("abort", onAbort, { once: true });

      record.promise.then(
        (bitmap) => {
          signal.removeEventListener("abort", onAbort);
          release();
          resolve(bitmap);
        },
        (err) => {
          signal.removeEventListener("abort", onAbort);
          release();
          reject(err);
        },
      );
    });
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
    if (typeof createImageBitmap !== "function") {
      return Promise.reject(new Error("createImageBitmap is not available in this environment"));
    }

    try {
      const buffer = new ArrayBuffer(entry.bytes.byteLength);
      new Uint8Array(buffer).set(entry.bytes);
      const blob = new Blob([buffer], { type: entry.mimeType });
      // `createImageBitmap` should always return a promise, but tests and
      // polyfills are not always well-behaved. Normalize to a promise so our
      // callers and internal bookkeeping remain predictable.
      return Promise.resolve(createImageBitmap(blob));
    } catch (err) {
      return Promise.reject(err);
    }
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
