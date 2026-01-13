import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

afterEach(() => {
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
});

