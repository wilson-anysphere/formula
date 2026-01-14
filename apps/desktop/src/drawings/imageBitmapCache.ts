import type { ImageEntry } from "./types";
import { MAX_PNG_DIMENSION, MAX_PNG_PIXELS, readPngDimensions } from "./pngDimensions";

export interface ImageBitmapCacheOptions {
  /**
   * Maximum number of decoded bitmaps to keep in the in-memory LRU cache.
   *
   * A value of 0 disables bitmap caching (but still dedupes concurrent decodes).
   *
   * @defaultValue 128
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
   * Callbacks registered by `getOrRequest()` callers waiting for this decode to
   * settle.
   */
  onReady: Set<() => void>;
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

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return typeof (value as { then?: unknown } | null)?.then === "function";
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
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(max ?? 128);
    this.negativeCacheMs = options.negativeCacheMs ?? 0;
  }

  setMaxEntries(maxEntries: number): void {
    this.maxEntries = ImageBitmapCache.normalizeMaxEntries(maxEntries);
    this.evictIfNeeded();
  }

  /**
   * Eagerly decode an image into an ImageBitmap.
   *
   * This is intended for insert flows so the subsequent render can draw using an
   * already-decoded (or decoding) bitmap.
   */
  preload(entry: ImageEntry): Promise<ImageBitmap> {
    return this.get(entry);
  }

  get(entry: ImageEntry, opts: ImageBitmapCacheGetOptions = {}): Promise<ImageBitmap> {
    const id = entry.id;

    // Opportunistically prune expired negative-cache entries so we don't retain
    // large numbers of failures indefinitely (the expiry window is short, but a
    // workbook could reference many distinct broken images).
    let now: number | undefined;
    if (this.negativeCacheMs > 0 && this.negativeCache.size > 0) {
      now = Date.now();
      this.pruneExpiredNegativeCache(now);
    }

    const existing = this.entries.get(id);
    if (existing) {
      // Mark as most-recently-used.
      this.touch(id, existing);
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
      const t = now ?? Date.now();
      if (cachedFailure.expiresAt > t) {
        return Promise.reject(cachedFailure.error);
      }
      this.negativeCache.delete(id);
    }

    const promise = ImageBitmapCache.decode(entry);
    const record: CacheEntry = {
      promise,
      bitmap: undefined as ImageBitmap | undefined,
      onReady: new Set(),
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
        if (current !== record || current.promise !== promise) {
          // This decode is no longer the active request for the image id (e.g.
          // it was invalidated or superseded). If every tracked consumer has
          // already aborted and there were no untracked consumers, ensure we
          // don't leak the decoded bitmap.
          record.onReady.clear();
          if (!record.pinned && record.waiters === 0) {
            ImageBitmapCache.tryClose(bitmap);
          }
          return;
        }

        // If the decode finishes after all callers have aborted (and no
        // untracked waiters exist), drop it immediately to avoid caching a bitmap
        // nobody will use.
        if (!record.pinned && record.waiters === 0 && record.onReady.size === 0) {
          this.entries.delete(id);
          ImageBitmapCache.tryClose(bitmap);
          return;
        }

        // A max size of 0 means caching is disabled. Still dedupe the in-flight
        // request, but don't store the decoded bitmap (and do not close it,
        // since the caller owns the resolved value).
        if (this.maxEntries === 0) {
          this.entries.delete(id);
          this.fireReadyCallbacks(record);
          // With caching disabled, any `getOrRequest()` callers will never receive the bitmap value.
          // Close it unless there is an active `get()` consumer waiting on the same decode.
          if (!record.pinned && record.waiters === 0) {
            ImageBitmapCache.tryClose(bitmap);
          }
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
        this.touch(id, record);

        this.evictIfNeeded();
        this.fireReadyCallbacks(record);
      },
      (err) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) return;

        this.__testOnly_failCount++;
        if (this.negativeCacheMs > 0) {
          // Replace + touch for predictable iteration order (useful for pruning).
          this.negativeCache.delete(id);
          this.negativeCache.set(id, { error: err, expiresAt: Date.now() + this.negativeCacheMs });
        }

        this.entries.delete(id);
        this.fireReadyCallbacks(record);
      },
    );

    return this.wrapWithAbort(id, record, opts.signal, Boolean(opts.signal));
  }

  /**
   * Returns a cached bitmap synchronously, or starts an async decode and returns
   * `null` immediately.
   *
   * When the decode finishes (success or failure), `onReady` is invoked so the
   * caller can schedule a re-render.
   */
  getOrRequest(entry: ImageEntry, onReady: () => void): ImageBitmap | null {
    const id = entry.id;

    // Opportunistically prune expired negative-cache entries so we don't retain
    // large numbers of failures indefinitely (the expiry window is short, but a
    // workbook could reference many distinct broken images).
    let now: number | undefined;
    if (this.negativeCacheMs > 0 && this.negativeCache.size > 0) {
      now = Date.now();
      this.pruneExpiredNegativeCache(now);
    }

    const existing = this.entries.get(id);
    if (existing) {
      this.touch(id, existing);
      if (existing.bitmap) return existing.bitmap;
      existing.onReady.add(onReady);
      return null;
    }

    const cachedFailure = this.negativeCache.get(id);
    if (cachedFailure) {
      const t = now ?? Date.now();
      if (cachedFailure.expiresAt > t) {
        return null;
      }
      this.negativeCache.delete(id);
    }

    const promise = ImageBitmapCache.decode(entry);
    const record: CacheEntry = {
      promise,
      bitmap: undefined as ImageBitmap | undefined,
      onReady: new Set([onReady]),
      waiters: 0,
      pinned: false,
    };
    this.entries.set(id, record);

    void promise.then(
      (bitmap) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) {
          record.onReady.clear();
          if (!record.pinned && record.waiters === 0) {
            ImageBitmapCache.tryClose(bitmap);
          }
          return;
        }

        if (!record.pinned && record.waiters === 0 && record.onReady.size === 0) {
          this.entries.delete(id);
          ImageBitmapCache.tryClose(bitmap);
          return;
        }

        if (this.maxEntries === 0) {
          this.entries.delete(id);
          this.fireReadyCallbacks(record);
          // With caching disabled, `getOrRequest()` callers will never receive the bitmap value.
          // Close it unless there is an active `get()` consumer waiting on the same decode.
          if (!record.pinned && record.waiters === 0) {
            ImageBitmapCache.tryClose(bitmap);
          }
          return;
        }

        if (!record.bitmap) {
          record.bitmap = bitmap;
          this.decodedCount++;
        } else if (record.bitmap !== bitmap) {
          ImageBitmapCache.tryClose(record.bitmap);
          record.bitmap = bitmap;
        }

        this.negativeCache.delete(id);
        this.touch(id, record);
        this.evictIfNeeded();
        this.fireReadyCallbacks(record);
      },
      (err) => {
        const current = this.entries.get(id);
        if (current !== record || current.promise !== promise) return;

        this.__testOnly_failCount++;
        if (this.negativeCacheMs > 0) {
          // Replace + touch for predictable iteration order (useful for pruning).
          this.negativeCache.delete(id);
          this.negativeCache.set(id, { error: err, expiresAt: Date.now() + this.negativeCacheMs });
        }

        this.entries.delete(id);
        this.fireReadyCallbacks(record);
      },
    );

    return null;
  }

  /**
   * Drop a cache entry (and any negative-cache failure state) for an image id.
   *
   * Note: `invalidate()` drops the cache entry but does **not** cancel in-flight `get()` calls.
   * `createImageBitmap` itself is not abortable; if a caller is awaiting a decode that becomes
   * stale (e.g. bytes changed, workbook/overlay teardown), the caller should pass an
   * `AbortSignal` to `get()` and abort it before invalidating so any eventually-decoded
   * ImageBitmap can be deterministically closed.
   */
  invalidate(imageId: string): void {
    const existing = this.entries.get(imageId);
    if (!existing) {
      this.negativeCache.delete(imageId);
      return;
    }

    existing.onReady.clear();
    this.entries.delete(imageId);
    this.negativeCache.delete(imageId);
    if (existing.bitmap) {
      this.decodedCount--;
      ImageBitmapCache.tryClose(existing.bitmap);
    }
  }

  /**
   * Close and drop all cached decoded bitmaps (and any negative-cache failure state).
   *
   * Note: `clear()` cannot cancel in-flight decodes (see note on `invalidate()`), but it does
   * ensure any *already-decoded* cached bitmaps are closed immediately.
   */
  clear(): void {
    for (const entry of this.entries.values()) {
      entry.onReady.clear();
      if (entry.bitmap) ImageBitmapCache.tryClose(entry.bitmap);
    }
    this.entries.clear();
    this.decodedCount = 0;
    this.negativeCache.clear();
  }

  /**
   * Alias for `clear()` to match common teardown naming (`dispose`/`destroy`).
   */
  dispose(): void {
    this.clear();
  }

  private touch(imageId: string, entry: CacheEntry): void {
    this.entries.delete(imageId);
    this.entries.set(imageId, entry);
  }

  private fireReadyCallbacks(entry: CacheEntry): void {
    if (entry.onReady.size === 0) return;
    const callbacks = Array.from(entry.onReady);
    entry.onReady.clear();
    for (const cb of callbacks) {
      try {
        const result = cb() as unknown;
        // Callbacks are expected to be synchronous, but unit tests sometimes stub them with async
        // mocks (returning a Promise). Swallow async rejections so decode completion does not
        // surface as an unhandled promise rejection.
        if (isThenable(result)) {
          void Promise.resolve(result).catch(() => {
            // Ignore async callback failures; cache bookkeeping must remain robust.
          });
        }
      } catch {
        // Ignore errors from caller-provided callbacks so cache bookkeeping stays robust.
      }
    }
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

  private pruneExpiredNegativeCache(now: number): void {
    // Entries are inserted in chronological order (we delete+set on refresh),
    // which means we can stop once we hit the first unexpired value.
    for (const [id, entry] of this.negativeCache) {
      if (entry.expiresAt > now) break;
      this.negativeCache.delete(id);
    }
  }

  private static normalizeMaxEntries(value: number): number {
    if (!Number.isFinite(value)) return 128;
    // Clamp to a small-ish non-negative integer to avoid pathological behavior.
    return Math.max(0, Math.floor(value));
  }

  private static async decode(entry: ImageEntry): Promise<ImageBitmap> {
    if (typeof createImageBitmap !== "function") {
      throw new Error("createImageBitmap is not available in this environment");
    }

    const dims = readPngDimensions(entry.bytes);
    if (dims) {
      const pixels = dims.width * dims.height;
      if (
        dims.width > MAX_PNG_DIMENSION ||
        dims.height > MAX_PNG_DIMENSION ||
        pixels > MAX_PNG_PIXELS
      ) {
        throw new Error(`Image dimensions too large (${dims.width}x${dims.height})`);
      }

      // Some callers/tests use a PNG header stub (signature + partial IHDR) to communicate dimensions
      // without including a full PNG payload. A valid PNG requires at least the full IHDR chunk
      // header + 13-byte IHDR data + CRC:
      //   signature(8) + length(4) + type(4) + data(13) + crc(4) = 33 bytes
      //
      // If we don't have that, decoding is guaranteed to fail; avoid invoking `createImageBitmap`
      // in that case so callers can rely on the "no decode attempted" invariant.
      if (entry.bytes.byteLength < 33) {
        throw new Error("Invalid PNG: truncated IHDR");
      }
    }

    // Blob construction already clones ArrayBufferView bytes; avoid a second manual copy.
    const blob = new Blob([entry.bytes], { type: entry.mimeType });

    try {
      // `createImageBitmap` should always return a promise, but tests and
      // polyfills are not always well-behaved. Normalize to a promise so our
      // callers and internal bookkeeping remain predictable.
      return await Promise.resolve(createImageBitmap(blob));
    } catch (err) {
      // Chrome can successfully decode certain malformed PNGs via `<img>`, but
      // rejects them via `createImageBitmap(blob)` (e.g. our real Excel fixtures).
      // Fall back to decoding through an `<img>` + canvas when we detect this
      // class of decode failure.
      const name = (err as any)?.name;
      if (name !== "InvalidStateError") throw err;

      const fallback = await ImageBitmapCache.decodeViaImageElement(blob).catch(() => null);
      if (fallback) return fallback;
      throw err;
    }
  }

  private static async decodeViaImageElement(blob: Blob): Promise<ImageBitmap> {
    // This fallback requires DOM APIs; if we're in a non-DOM environment, just
    // fail and let the caller surface the original decode error.
    if (typeof document === "undefined") {
      throw new Error("Image decode fallback requires DOM APIs (missing document)");
    }
    if (typeof Image !== "function") {
      throw new Error("Image decode fallback requires DOM APIs (missing Image)");
    }
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") {
      throw new Error("Image decode fallback requires URL.createObjectURL");
    }

    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      await new Promise<void>((resolve, reject) => {
        const timeoutMs = 5_000;
        let timeoutId: ReturnType<typeof setTimeout> | null = null;
        let settled = false;

        const finish = (fn: () => void) => {
          if (settled) return;
          settled = true;
          if (timeoutId !== null) {
            try {
              clearTimeout(timeoutId);
            } catch {
              // Ignore clear failures (best-effort).
            }
            timeoutId = null;
          }
          // Drop references to callbacks to allow GC even if the image object remains alive.
          img.onload = null;
          img.onerror = null;
          fn();
        };

        img.onload = () => {
          finish(resolve);
        };
        img.onerror = () => {
          finish(() => reject(new Error("Image decode fallback failed to load <img>")));
        };

        if (typeof setTimeout === "function") {
          timeoutId = setTimeout(() => {
            timeoutId = null;
            finish(() => reject(new Error("Image decode fallback timed out")));
          }, timeoutMs);
        }

        // Assign the src after wiring handlers so we don't miss synchronous load events in tests/polyfills.
        img.src = url;
      });

      const canvas = document.createElement("canvas");
      const width = (img as any).naturalWidth ?? img.width;
      const height = (img as any).naturalHeight ?? img.height;
      canvas.width = Number.isFinite(width) && width > 0 ? width : 1;
      canvas.height = Number.isFinite(height) && height > 0 ? height : 1;
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Image decode fallback missing 2D canvas context");
      ctx.drawImage(img, 0, 0);

      // `createImageBitmap(img)` can still fail for the same malformed inputs even
      // if the image element decodes successfully. Converting through an
      // intermediate canvas is more robust.
      return await Promise.resolve(createImageBitmap(canvas));
    } finally {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // Ignore revoke failures (best-effort).
      }
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
