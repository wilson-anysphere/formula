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
});
