import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
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

  it("removes a failed decode from the cache and allows a subsequent retry", async () => {
    const cache = new ImageBitmapCache({ negativeCacheMs: 0 });
    const entry = makeEntry();

    const err = new Error("bad bytes");
    const bitmap = {} as ImageBitmap;

    const createImageBitmapMock = vi
      .fn()
      .mockRejectedValueOnce(err)
      .mockResolvedValueOnce(bitmap);
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

    // Let the underlying decode complete to ensure there are no unhandled
    // promise rejections during the test run.
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

    // When the decode eventually resolves, the bitmap should be closed since no
    // one is still awaiting it.
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
