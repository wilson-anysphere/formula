import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

function createEntry(id: string): ImageEntry {
  return { id, bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
}

function createPngHeaderBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  // 13-byte IHDR chunk length.
  bytes[8] = 0x00;
  bytes[9] = 0x00;
  bytes[10] = 0x00;
  bytes[11] = 0x0d;
  // IHDR chunk type.
  bytes[12] = 0x49;
  bytes[13] = 0x48;
  bytes[14] = 0x44;
  bytes[15] = 0x52;

  const view = new DataView(bytes.buffer);
  view.setUint32(16, width, false);
  view.setUint32(20, height, false);
  return bytes;
}

function createJpegHeaderBytes(width: number, height: number): Uint8Array {
  // Minimal structure: SOI + APP0 (dummy) + SOF0 with width/height.
  // This is not a complete/valid JPEG, but it includes enough header structure for our
  // dimension parser to extract the advertised size.
  const bytes = new Uint8Array(33);
  let o = 0;
  // SOI
  bytes[o++] = 0xff;
  bytes[o++] = 0xd8;
  // APP0 marker
  bytes[o++] = 0xff;
  bytes[o++] = 0xe0;
  // APP0 length: 16 bytes (includes these 2 length bytes) + 14 bytes payload.
  bytes[o++] = 0x00;
  bytes[o++] = 0x10;
  o += 14;
  // SOF0 marker
  bytes[o++] = 0xff;
  bytes[o++] = 0xc0;
  // SOF0 length: 11 bytes (includes these 2 length bytes) = 9 bytes payload.
  bytes[o++] = 0x00;
  bytes[o++] = 0x0b;
  // Precision
  bytes[o++] = 0x08;
  // Height (big-endian)
  bytes[o++] = (height >> 8) & 0xff;
  bytes[o++] = height & 0xff;
  // Width (big-endian)
  bytes[o++] = (width >> 8) & 0xff;
  bytes[o++] = width & 0xff;
  // Components (1) + component spec (3 bytes)
  bytes[o++] = 0x01;
  bytes[o++] = 0x01;
  bytes[o++] = 0x11;
  bytes[o++] = 0x00;
  return bytes;
}

function createGifHeaderBytes(width: number, height: number): Uint8Array {
  // GIF header (GIF89a) + logical screen width/height (little-endian).
  const bytes = new Uint8Array(10);
  bytes[0] = 0x47; // G
  bytes[1] = 0x49; // I
  bytes[2] = 0x46; // F
  bytes[3] = 0x38; // 8
  bytes[4] = 0x39; // 9
  bytes[5] = 0x61; // a
  bytes[6] = width & 0xff;
  bytes[7] = (width >> 8) & 0xff;
  bytes[8] = height & 0xff;
  bytes[9] = (height >> 8) & 0xff;
  return bytes;
}

function createWebpVp8xHeaderBytes(width: number, height: number): Uint8Array {
  // Minimal structure:
  //  - RIFF header + WEBP signature
  //  - VP8X chunk type
  //  - width/height minus one (24-bit little-endian)
  const bytes = new Uint8Array(30);
  bytes[0] = 0x52; // R
  bytes[1] = 0x49; // I
  bytes[2] = 0x46; // F
  bytes[3] = 0x46; // F
  bytes[8] = 0x57; // W
  bytes[9] = 0x45; // E
  bytes[10] = 0x42; // B
  bytes[11] = 0x50; // P
  bytes[12] = 0x56; // V
  bytes[13] = 0x50; // P
  bytes[14] = 0x38; // 8
  bytes[15] = 0x58; // X

  const w = Math.max(1, Math.floor(width)) - 1;
  const h = Math.max(1, Math.floor(height)) - 1;
  bytes[24] = w & 0xff;
  bytes[25] = (w >> 8) & 0xff;
  bytes[26] = (w >> 16) & 0xff;
  bytes[27] = h & 0xff;
  bytes[28] = (h >> 8) & 0xff;
  bytes[29] = (h >> 16) & 0xff;
  return bytes;
}

function createBmpHeaderBytes(width: number, height: number): Uint8Array {
  // Minimal structure: BMP file header + BITMAPINFOHEADER with width/height.
  // This is not a complete/valid BMP, but it includes enough header structure for our
  // dimension parser to extract the advertised size.
  const bytes = new Uint8Array(54);
  const view = new DataView(bytes.buffer);

  // Signature "BM"
  bytes[0] = 0x42;
  bytes[1] = 0x4d;
  // Pixel array offset (header size)
  view.setUint32(10, 54, true);
  // DIB header size (BITMAPINFOHEADER)
  view.setUint32(14, 40, true);
  // Width/height (signed 32-bit)
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  // Planes
  view.setUint16(26, 1, true);
  // Bits per pixel
  view.setUint16(28, 24, true);
  return bytes;
}

function createSvgBytes(svg: string): Uint8Array {
  try {
    if (typeof TextEncoder !== "undefined") {
      return new TextEncoder().encode(svg);
    }
  } catch {
    // Fall through to manual encoding.
  }

  // ASCII-only fallback (our test fixtures don't require full UTF-8 support).
  const bytes = new Uint8Array(svg.length);
  for (let i = 0; i < svg.length; i += 1) bytes[i] = svg.charCodeAt(i) & 0xff;
  return bytes;
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

  it("rejects PNG images with huge IHDR dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry: ImageEntry = { id: "png_bomb", bytes: createPngHeaderBytes(10_001, 1), mimeType: "image/png" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects JPEG images with huge dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry: ImageEntry = { id: "jpeg_bomb", bytes: createJpegHeaderBytes(10_001, 1), mimeType: "image/jpeg" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects GIF images with huge dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry: ImageEntry = { id: "gif_bomb", bytes: createGifHeaderBytes(10_001, 1), mimeType: "image/gif" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects WebP images with huge dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry: ImageEntry = { id: "webp_bomb", bytes: createWebpVp8xHeaderBytes(10_001, 1), mimeType: "image/webp" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects BMP images with huge dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry: ImageEntry = { id: "bmp_bomb", bytes: createBmpHeaderBytes(10_001, 1), mimeType: "image/bmp" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects SVG images with huge dimensions without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="10001" height="1"></svg>`;
    const entry: ImageEntry = { id: "svg_bomb", bytes: createSvgBytes(svg), mimeType: "image/svg+xml" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects PNG images that exceed the pixel limit without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    // 9000 x 9000 = 81M pixels (over the 50M limit, but each dimension is < 10k).
    const entry: ImageEntry = { id: "png_pixel_bomb", bytes: createPngHeaderBytes(9000, 9000), mimeType: "image/png" };

    await expect(cache.get(entry)).rejects.toThrow(/Image dimensions too large/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rejects PNG images with truncated IHDR without invoking createImageBitmap", async () => {
    const createImageBitmapMock = vi.fn(() => Promise.resolve({ close: vi.fn() } as unknown as ImageBitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    // Width/height are valid, but the byte payload is too short to include a full IHDR chunk + CRC.
    const entry: ImageEntry = { id: "png_truncated_ihdr", bytes: createPngHeaderBytes(1, 1), mimeType: "image/png" };

    await expect(cache.get(entry)).rejects.toThrow(/Invalid PNG: truncated IHDR/);
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("rethrows InvalidStateError when the DOM fallback is unavailable (node env)", async () => {
    // Ensure this test actually runs in a non-DOM environment. (The ImageBitmapCache fallback path
    // is DOM-only and should be a no-op under Node.)
    expect(typeof document).toBe("undefined");

    const err = new Error("decode failed");
    (err as any).name = "InvalidStateError";
    const createImageBitmapMock = vi.fn(() => Promise.reject(err));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const cache = new ImageBitmapCache({ maxEntries: 10 });
    const entry = createEntry("img_invalid_state");

    await expect(cache.get(entry)).rejects.toMatchObject({ name: "InvalidStateError" });
    // The DOM fallback should be skipped (document is missing), so we only attempt the blob decode once.
    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
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
