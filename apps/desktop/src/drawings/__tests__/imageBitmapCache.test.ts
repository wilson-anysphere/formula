import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function makeEntry(id = "img_1"): ImageEntry {
  return {
    id,
    bytes: new Uint8Array([1, 2, 3, 4]),
    mimeType: "image/png",
  };
}

describe("ImageBitmapCache", () => {
  it("clears the negative cache on invalidate so callers can retry immediately", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const cache = new ImageBitmapCache({ negativeCacheMs: 10_000 });
    const entry = makeEntry();

    const err = new Error("decode failed");
    const bitmap = {} as ImageBitmap;

    const createImageBitmapMock = vi.fn().mockRejectedValueOnce(err).mockResolvedValueOnce(bitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    await expect(cache.get(entry)).rejects.toBe(err);
    // Ensure the cache's internal bookkeeping has a chance to run.
    await Promise.resolve();

    // Negative cache should prevent tight retry loops (no new decode attempt yet).
    await expect(cache.get(entry)).rejects.toBe(err);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

    // When the bytes are replaced, callers typically invalidate the cache entry.
    cache.invalidate(entry.id);

    await expect(cache.get(entry)).resolves.toBe(bitmap);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
  });

  it("dedupes concurrent preload calls", async () => {
    let resolveBitmap: ((value: ImageBitmap) => void) | null = null;

    const createImageBitmapMock = vi.fn(
      () =>
        new Promise<ImageBitmap>((resolve) => {
          resolveBitmap = resolve;
        }),
    );
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 16, negativeCacheMs: 0 });
    const img = makeEntry("img_1");

    const p1 = cache.preload(img);
    const p2 = cache.preload(img);

    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
    // The cache should return the same in-flight promise for a given image id.
    expect(p2).toBe(p1);

    const bitmap = { id: "decoded" } as any as ImageBitmap;
    resolveBitmap?.(bitmap);

    const [b1, b2] = await Promise.all([p1, p2]);
    expect(b1).toBe(bitmap);
    expect(b2).toBe(bitmap);
  });

  it("evicts least-recently-used resolved bitmaps when over the limit", async () => {
    let decodeCount = 0;
    const createImageBitmapMock = vi.fn(async () => ({ id: `bitmap_${decodeCount++}` } as any as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 2, negativeCacheMs: 0 });
    const a = makeEntry("a");
    const b = makeEntry("b");
    const c = makeEntry("c");

    // Fill the cache with a and b.
    await cache.get(a);
    await cache.get(b);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);

    // Touch a so b becomes the LRU entry.
    await cache.get(a);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);

    // Inserting c should evict b.
    await cache.get(c);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(3);

    // a is still cached.
    await cache.get(a);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(3);

    // b was evicted, so fetching it again should trigger another decode.
    await cache.get(b);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(4);
  });

  it("removes a failed decode from the cache and allows a subsequent retry", async () => {
    const cache = new ImageBitmapCache({ negativeCacheMs: 0 });
    const entry = makeEntry();

    const err = new Error("bad bytes");
    const bitmap = {} as ImageBitmap;

    const createImageBitmapMock = vi.fn().mockRejectedValueOnce(err).mockResolvedValueOnce(bitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    await expect(cache.get(entry)).rejects.toBe(err);
    expect(cache.__testOnly_failCount).toBe(1);

    await expect(cache.get(entry)).resolves.toBe(bitmap);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
  });

  it("supports aborting an inflight decode and cleans up the cache entry", async () => {
    const cache = new ImageBitmapCache({ negativeCacheMs: 0 });
    const entry = makeEntry();

    let resolveDecode!: (bitmap: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });

    const createImageBitmapMock = vi.fn().mockReturnValueOnce(inflightDecode).mockResolvedValueOnce({} as ImageBitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const controller = new AbortController();
    const inflight = cache.get(entry, { signal: controller.signal });
    controller.abort();

    await expect(inflight).rejects.toMatchObject({ name: "AbortError" });

    // Let the underlying decode complete to ensure there are no unhandled promise rejections.
    resolveDecode({} as ImageBitmap);
    await Promise.resolve();

    await cache.get(entry);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
  });

  it("does not drop a shared in-flight decode when one waiter aborts", async () => {
    const cache = new ImageBitmapCache({ negativeCacheMs: 0 });
    const entry = makeEntry();

    let resolveDecode!: (bitmap: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });
    const bitmap = {} as ImageBitmap;

    const createImageBitmapMock = vi.fn().mockReturnValue(inflightDecode);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const controller1 = new AbortController();
    const controller2 = new AbortController();

    const p1 = cache.get(entry, { signal: controller1.signal });
    const p2 = cache.get(entry, { signal: controller2.signal });

    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

    controller1.abort();
    await expect(p1).rejects.toMatchObject({ name: "AbortError" });

    resolveDecode(bitmap);
    await expect(p2).resolves.toBe(bitmap);

    // The decode should not have been restarted.
    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

    // And the bitmap should remain available from the cache.
    await expect(cache.get(entry)).resolves.toBe(bitmap);
  });

  it("closes decoded bitmaps from stale in-flight decodes when all waiters abort and the entry is invalidated", async () => {
    const cache = new ImageBitmapCache({ negativeCacheMs: 0 });
    const entry = makeEntry();

    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;

    let resolveDecode!: (value: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });

    vi.stubGlobal("createImageBitmap", vi.fn(() => inflightDecode) as unknown as typeof createImageBitmap);

    const controller = new AbortController();
    const p = cache.get(entry, { signal: controller.signal });
    controller.abort();
    await expect(p).rejects.toMatchObject({ name: "AbortError" });

    // Drop the cache entry while the decode is still in-flight.
    cache.invalidate(entry.id);

    // When the decode eventually resolves, the bitmap should be closed since no one is still awaiting it.
    resolveDecode(bitmap);
    await Promise.resolve();

    expect(close).toHaveBeenCalledTimes(1);
  });

  it("honors negativeCacheMs by suppressing immediate retries after a failure", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    try {
      const cache = new ImageBitmapCache({ negativeCacheMs: 250 });
      const entry = makeEntry();

      const err = new Error("bad bytes");
      const bitmap = {} as ImageBitmap;

      const createImageBitmapMock = vi.fn().mockRejectedValueOnce(err).mockResolvedValueOnce(bitmap);
      vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

      await expect(cache.get(entry)).rejects.toBe(err);
      expect(cache.__testOnly_failCount).toBe(1);
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

      // Within the negative cache window, we should not attempt another decode.
      await expect(cache.get(entry)).rejects.toBe(err);
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

      // After expiry, retry should be allowed.
      vi.advanceTimersByTime(300);
      vi.setSystemTime(300);
      await expect(cache.get(entry)).resolves.toBe(bitmap);
      expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });
});
