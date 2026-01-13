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
    await cache.get(b);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(4);
  });
});

