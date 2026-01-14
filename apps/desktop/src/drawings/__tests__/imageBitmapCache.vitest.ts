import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

function createEntry(id: string): ImageEntry {
  return { id, bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
}

describe("ImageBitmapCache", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("dedupes concurrent decode requests for the same id", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;

    let resolve!: (value: ImageBitmap) => void;
    const decodePromise = new Promise<ImageBitmap>((res) => {
      resolve = res;
    });

    const createImageBitmapMock = vi.fn(() => decodePromise);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_1");

    const p1 = cache.get(entry);
    const p2 = cache.get(entry);

    expect(p1).toBe(p2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

    resolve(bitmap);

    await expect(p1).resolves.toBe(bitmap);
    await expect(p2).resolves.toBe(bitmap);
    expect(close).not.toHaveBeenCalled();
  });

  it("evicts least-recently-used entries and closes evicted bitmaps", async () => {
    const bitmaps: Array<{ id: string; close: ReturnType<typeof vi.fn> }> = [];
    const createImageBitmapMock = vi.fn(() => {
      const idx = bitmaps.length + 1;
      const close = vi.fn();
      const bitmap = { close } as unknown as ImageBitmap;
      bitmaps.push({ id: `bitmap_${idx}`, close });
      return Promise.resolve(bitmap);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 2 });
    const a = createEntry("a");
    const b = createEntry("b");
    const c = createEntry("c");

    const bitmapA = await cache.get(a);
    const bitmapB = await cache.get(b);

    // Touch `a` so `b` becomes least-recently-used.
    await cache.get(a);

    const bitmapC = await cache.get(c);

    expect(bitmapA).toBeDefined();
    expect(bitmapB).toBeDefined();
    expect(bitmapC).toBeDefined();

    // `b` should have been evicted + closed.
    expect((bitmapB as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapA as any).close).not.toHaveBeenCalled();
    expect((bitmapC as any).close).not.toHaveBeenCalled();

    // Re-requesting `b` should trigger a new decode.
    const bitmapB2 = await cache.get(b);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(4);

    // Clearing should close remaining cached bitmaps, without double-closing the
    // already-evicted bitmap.
    cache.clear();
    expect((bitmapB as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapA as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapC as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapB2 as any).close).toHaveBeenCalledTimes(1);
  });

  it("does not repopulate the cache from a stale in-flight decode after invalidate()", async () => {
    const bitmap1 = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmap2 = { close: vi.fn() } as unknown as ImageBitmap;

    let resolve1!: (value: ImageBitmap) => void;
    const p1 = new Promise<ImageBitmap>((res) => {
      resolve1 = res;
    });

    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => p1)
      .mockImplementationOnce(() => Promise.resolve(bitmap2));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_1");

    const first = cache.get(entry);
    cache.invalidate(entry.id);
    resolve1(bitmap1);
    await expect(first).resolves.toBe(bitmap1);

    // Should trigger a new decode since the stale one was invalidated.
    await cache.get(entry);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
  });

  it("invalidate() closes cached bitmaps and forces a re-decode", async () => {
    const bitmap1 = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmap2 = { close: vi.fn() } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => Promise.resolve(bitmap1))
      .mockImplementationOnce(() => Promise.resolve(bitmap2));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_1");

    await expect(cache.get(entry)).resolves.toBe(bitmap1);
    cache.invalidate(entry.id);
    expect((bitmap1 as any).close).toHaveBeenCalledTimes(1);

    await expect(cache.get(entry)).resolves.toBe(bitmap2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
    expect((bitmap2 as any).close).not.toHaveBeenCalled();
  });

  it("clear() closes all cached bitmaps", async () => {
    const bitmapA = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmapB = { close: vi.fn() } as unknown as ImageBitmap;
    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => Promise.resolve(bitmapA))
      .mockImplementationOnce(() => Promise.resolve(bitmapB));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    await cache.get(createEntry("a"));
    await cache.get(createEntry("b"));

    cache.clear();

    expect((bitmapA as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapB as any).close).toHaveBeenCalledTimes(1);
  });

  it("does not immediately evict a late-resolving decode before callers can use it", async () => {
    const bitmapA = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmapB = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmapB2 = { close: vi.fn() } as unknown as ImageBitmap;

    let resolveA!: (value: ImageBitmap) => void;
    let resolveB!: (value: ImageBitmap) => void;
    const pA = new Promise<ImageBitmap>((res) => {
      resolveA = res;
    });
    const pB = new Promise<ImageBitmap>((res) => {
      resolveB = res;
    });

    const createImageBitmapMock = vi
      .fn()
      .mockImplementationOnce(() => pA)
      .mockImplementationOnce(() => pB)
      .mockImplementationOnce(() => Promise.resolve(bitmapB2))
      // Subsequent decodes should return a promise so `cache.get()` never tries to
      // attach handlers to `undefined` when we re-request after eviction.
      .mockImplementation(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 1 });

    const promiseA = cache.get(createEntry("a"));
    const promiseB = cache.get(createEntry("b"));

    // Resolve out-of-order: B first, then A.
    resolveB(bitmapB);
    await expect(promiseB).resolves.toBe(bitmapB);

    resolveA(bitmapA);
    await expect(promiseA).resolves.toBe(bitmapA);

    // The cache should not evict+close the bitmap in the same microtask that
    // resolves its promise (the internal `.then` handlers run before caller
    // `await`/`.then` handlers attached later).
    //
    // With maxEntries=1, `b` should be the one evicted+closed once `a` also resolves.
    expect((bitmapA as any).close).not.toHaveBeenCalled();
    expect((bitmapB as any).close).toHaveBeenCalledTimes(1);

    // `a` should still be cached; `b` should require a new decode.
    await expect(cache.get(createEntry("a"))).resolves.toBe(bitmapA);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
    await cache.get(createEntry("b"));
    expect(createImageBitmapMock).toHaveBeenCalledTimes(3);
  });

  it("setMaxEntries() evicts down to the new limit and closes evicted bitmaps", async () => {
    const bitmapA = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmapB = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmapC = { close: vi.fn() } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => Promise.resolve(bitmapA))
      .mockImplementationOnce(() => Promise.resolve(bitmapB))
      .mockImplementationOnce(() => Promise.resolve(bitmapC));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 3 });
    await cache.get(createEntry("a"));
    await cache.get(createEntry("b"));
    await cache.get(createEntry("c"));

    // Make `b` the most recently used so `a` is the LRU.
    await cache.get(createEntry("b"));

    cache.setMaxEntries(1);

    // Should keep only `b` and evict/close `a` + `c`.
    expect((bitmapB as any).close).not.toHaveBeenCalled();
    expect((bitmapA as any).close).toHaveBeenCalledTimes(1);
    expect((bitmapC as any).close).toHaveBeenCalledTimes(1);
  });

  it("supports disabling caching via maxEntries=0 (still dedupes concurrent requests)", async () => {
    const bitmap1 = { close: vi.fn() } as unknown as ImageBitmap;
    const bitmap2 = { close: vi.fn() } as unknown as ImageBitmap;

    let resolve!: (value: ImageBitmap) => void;
    const pending = new Promise<ImageBitmap>((res) => {
      resolve = res;
    });

    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => pending)
      .mockImplementationOnce(() => Promise.resolve(bitmap2));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 0 });
    const entry = createEntry("img_1");

    const p1 = cache.get(entry);
    const p2 = cache.get(entry);
    expect(p1).toBe(p2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

    resolve(bitmap1);
    await expect(p1).resolves.toBe(bitmap1);

    // Not cached: should re-decode on subsequent get.
    await expect(cache.get(entry)).resolves.toBe(bitmap2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);

    // When caching is disabled, the cache should not close bitmaps it doesn't retain.
    expect((bitmap1 as any).close).not.toHaveBeenCalled();
    expect((bitmap2 as any).close).not.toHaveBeenCalled();
  });

  it("allows retry after a failed decode", async () => {
    const bitmap = { close: vi.fn() } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn()
      .mockImplementationOnce(() => Promise.reject(new Error("decode failed")))
      .mockImplementationOnce(() => Promise.resolve(bitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_1");

    await expect(cache.get(entry)).rejects.toThrow("decode failed");
    await expect(cache.get(entry)).resolves.toBe(bitmap);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
  });

  it("does not start decoding when the AbortSignal is already aborted", async () => {
    const bitmap = { close: vi.fn() } as unknown as ImageBitmap;
    const createImageBitmapMock = vi.fn(() => Promise.resolve(bitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_1");

    const controller = new AbortController();
    controller.abort();

    await expect(cache.get(entry, { signal: controller.signal })).rejects.toMatchObject({ name: "AbortError" });
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("negativeCacheMs prevents tight retry loops after decode failures", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    try {
      const bitmap = { close: vi.fn() } as unknown as ImageBitmap;
      const createImageBitmapMock = vi
        .fn()
        .mockRejectedValueOnce(new Error("decode failed"))
        .mockResolvedValueOnce(bitmap);
      vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

      const cache = new ImageBitmapCache({ maxEntries: 10, negativeCacheMs: 250 });
      const entry = createEntry("img_1");

      await expect(cache.get(entry)).rejects.toThrow("decode failed");
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

      // Within the negative-cache window we should not invoke a new decode.
      await expect(cache.get(entry)).rejects.toThrow("decode failed");
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);

      // After expiry we should retry.
      vi.advanceTimersByTime(300);
      vi.setSystemTime(300);
      await expect(cache.get(entry)).resolves.toBe(bitmap);
      expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it("prunes expired negative cache entries so failure metadata cannot grow unbounded", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    try {
      const createImageBitmapMock = vi.fn().mockRejectedValue(new Error("decode failed"));
      vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

      const cache = new ImageBitmapCache({ maxEntries: 10, negativeCacheMs: 100 });

      await expect(cache.get(createEntry("a"))).rejects.toThrow("decode failed");
      await expect(cache.get(createEntry("b"))).rejects.toThrow("decode failed");

      const negativeCache = (cache as any).negativeCache as Map<string, unknown>;
      expect(negativeCache.size).toBe(2);

      // Move time forward past the expiry window and trigger a new `get()` call,
      // which should prune old failures.
      vi.advanceTimersByTime(200);
      vi.setSystemTime(200);

      await expect(cache.get(createEntry("c"))).rejects.toThrow("decode failed");
      // Flush the internal `.then` handlers that populate negative cache entries.
      await Promise.resolve();

      expect([...negativeCache.keys()]).toEqual(["c"]);
      expect(createImageBitmapMock).toHaveBeenCalledTimes(3);
    } finally {
      vi.useRealTimers();
    }
  });
});
