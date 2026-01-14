// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

type Recording = {
  drawImages: any[][];
  fillTexts: any[][];
};

function createRecordingContext(canvas: HTMLCanvasElement): { ctx: CanvasRenderingContext2D; rec: Recording } {
  const rec: Recording = {
    drawImages: [],
    fillTexts: []
  };

  let font = "";
  let fillStyle: string | CanvasGradient | CanvasPattern = "#000";
  let strokeStyle: string | CanvasGradient | CanvasPattern = "#000";
  let lineWidth = 1;
  let textAlign: CanvasTextAlign = "left";
  let textBaseline: CanvasTextBaseline = "alphabetic";
  let globalAlpha = 1;
  let imageSmoothingEnabled = false;

  const ctx: Partial<CanvasRenderingContext2D> = {
    canvas,
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
    },
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value: string | CanvasGradient | CanvasPattern) {
      fillStyle = value;
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value: string | CanvasGradient | CanvasPattern) {
      strokeStyle = value;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(value: number) {
      lineWidth = value;
    },
    get textAlign() {
      return textAlign;
    },
    set textAlign(value: CanvasTextAlign) {
      textAlign = value;
    },
    get textBaseline() {
      return textBaseline;
    },
    set textBaseline(value: CanvasTextBaseline) {
      textBaseline = value;
    },
    get globalAlpha() {
      return globalAlpha;
    },
    set globalAlpha(value: number) {
      globalAlpha = value;
    },
    get imageSmoothingEnabled() {
      return imageSmoothingEnabled;
    },
    set imageSmoothingEnabled(value: boolean) {
      imageSmoothingEnabled = value;
    },

    setTransform: vi.fn(),
    clearRect: vi.fn(),
    fillRect: vi.fn(),
    strokeRect: vi.fn(),
    beginPath: vi.fn(),
    rect: vi.fn(),
    clip: vi.fn(),
    fill: vi.fn(),
    stroke: vi.fn(),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    closePath: vi.fn(),
    save: vi.fn(),
    restore: vi.fn(),
    translate: vi.fn(),
    rotate: vi.fn(),
    setLineDash: vi.fn(),
    createPattern: vi.fn(() => ({}) as any),

    drawImage: vi.fn((...args: any[]) => {
      rec.drawImages.push(args);
    }),
    fillText: vi.fn((...args: any[]) => {
      rec.fillTexts.push(args);
    }),
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  };

  return { ctx: ctx as unknown as CanvasRenderingContext2D, rec };
}

function createPngHeaderBytes(width: number, height: number, totalBytes = 33): Uint8Array {
  const bytes = new Uint8Array(totalBytes);
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
      const encoded = new TextEncoder().encode(svg);
      // Ensure the returned view uses the current realm's Uint8Array constructor.
      return new Uint8Array(encoded);
    }
  } catch {
    // Fall through to manual encoding.
  }

  // ASCII-only fallback (our test fixtures don't require full UTF-8 support).
  const bytes = new Uint8Array(svg.length);
  for (let i = 0; i < svg.length; i += 1) bytes[i] = svg.charCodeAt(i) & 0xff;
  return bytes;
}

async function flushMicrotasks(): Promise<void> {
  // Several turns helps flush async chains that include multiple `await` boundaries
  // (resolver -> header sniff -> createImageBitmap -> finally handlers).
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
}

describe("CanvasGridRenderer image cells", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  const installContexts = (
    contexts: Map<HTMLCanvasElement, CanvasRenderingContext2D>,
  ): void => {
    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  };

  it("draws image thumbnails with the expected destination rect", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", width: 100, height: 50 }
            }
          : null
    };

    const bitmap = { width: 200, height: 100 } as any;
    vi.stubGlobal("createImageBitmap", vi.fn(async () => bitmap));

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(content.rec.drawImages.length).toBeGreaterThan(0);

    const args = content.rec.drawImages[content.rec.drawImages.length - 1]!;
    // drawImage(image, dx, dy, dw, dh)
    expect(args.length).toBeGreaterThanOrEqual(5);
    expect(args[1]).toBeCloseTo(4);
    expect(args[2]).toBeCloseTo(2);
    expect(args[3]).toBeCloseTo(92);
    expect(args[4]).toBeCloseTo(46);
  });

  it("renders a placeholder (no drawImage) when the resolver returns missing", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "missing", altText: "Missing image", width: 100, height: 50 }
            }
          : null
    };

    const imageResolver = vi.fn(async () => null);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Missing image")).toBe(true);
  });

  it("rejects PNG images with huge IHDR dimensions without invoking createImageBitmap", async () => {
    // Avoid synchronous RAF side effects; this test controls when renders occur.
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_bomb", altText: "Bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createPngHeaderBytes(10_001, 1);
    // Return raw bytes so the renderer's PNG guard can run in jsdom environments (jsdom's Blob
    // implementation does not support reading bytes back via `arrayBuffer`).
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Bomb")).toBe(true);
  });

  it("rejects PNG Blob images with huge IHDR dimensions without invoking createImageBitmap", async () => {
    // Avoid synchronous RAF side effects; this test controls when renders occur.
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_bomb_blob", altText: "Blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createPngHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Blob bomb")).toBe(true);
  });

  it("rejects JPEG images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "jpeg_bomb", altText: "JPEG bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createJpegHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("jpeg_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "JPEG bomb")).toBe(true);
  });

  it("rejects JPEG Blob images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "jpeg_bomb_blob", altText: "JPEG blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createJpegHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/jpeg" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("jpeg_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "JPEG blob bomb")).toBe(true);
  });

  it("rejects GIF images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "gif_bomb", altText: "GIF bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // GIF dimensions are 16-bit; exceed the shared MAX_PNG_DIMENSION (10k).
    const bytes = createGifHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("gif_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "GIF bomb")).toBe(true);
  });

  it("rejects GIF Blob images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "gif_bomb_blob", altText: "GIF blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createGifHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/gif" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("gif_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "GIF blob bomb")).toBe(true);
  });

  it("rejects WebP images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "webp_bomb", altText: "WebP bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createWebpVp8xHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("webp_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "WebP bomb")).toBe(true);
  });

  it("rejects WebP Blob images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "webp_bomb_blob", altText: "WebP blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createWebpVp8xHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/webp" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("webp_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "WebP blob bomb")).toBe(true);
  });

  it("rejects BMP images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "bmp_bomb", altText: "BMP bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createBmpHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("bmp_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "BMP bomb")).toBe(true);
  });

  it("rejects BMP Blob images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "bmp_bomb_blob", altText: "BMP blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const bytes = createBmpHeaderBytes(10_001, 1);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/bmp" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("bmp_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "BMP blob bomb")).toBe(true);
  });

  it("rejects SVG images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "svg_bomb", altText: "SVG bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="10001" height="1"></svg>`;
    const imageResolver = vi.fn(async () => createSvgBytes(svg));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("svg_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "SVG bomb")).toBe(true);
  });

  it("rejects SVG Blob images with huge dimensions without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "svg_bomb_blob", altText: "SVG blob bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // Ensure `<svg ...>` appears after the initial TYPE_SNIFF_BYTES (32) to exercise the
    // larger SVG header read path in guardPngBlob.
    const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="10001" height="1"></svg>`;
    const bytes = createSvgBytes(svg);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/svg+xml" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("svg_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "SVG blob bomb")).toBe(true);
  });

  it("rejects PNG Blob images that exceed the pixel limit without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_pixel_bomb_blob", altText: "Blob pixel bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // 9000 x 9000 = 81M pixels (over the 50M limit, but each dimension is < 10k).
    const bytes = createPngHeaderBytes(9000, 9000);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_pixel_bomb_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Blob pixel bomb")).toBe(true);
  });

  it("rejects PNG Blob images with truncated IHDR without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_truncated_ihdr_blob", altText: "Blob truncated", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // Valid IHDR dimensions, but too few bytes for IHDR chunk + CRC.
    const bytes = createPngHeaderBytes(1, 1, 24);
    const imageResolver = vi.fn(async () => new Blob([bytes], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_truncated_ihdr_blob") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Blob truncated")).toBe(true);
  });

  it("rejects PNG images that exceed the pixel limit without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_pixel_bomb", altText: "Pixel bomb", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // 9000 x 9000 = 81M pixels (over the 50M limit, but each dimension is < 10k).
    const bytes = createPngHeaderBytes(9000, 9000);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_pixel_bomb") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Pixel bomb")).toBe(true);
  });

  it("rejects PNG images with truncated IHDR without invoking createImageBitmap", async () => {
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "png_truncated_ihdr", altText: "Truncated", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn(async () => ({ width: 10, height: 10 } as any));
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    // Valid IHDR dimensions, but too few bytes for IHDR chunk + CRC.
    const bytes = createPngHeaderBytes(1, 1, 24);
    const imageResolver = vi.fn(async () => bytes);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    const pending = (renderer as any).imageBitmapCache.get("png_truncated_ihdr") as
      | { state: "pending"; promise: Promise<void> }
      | { state: string };
    if (pending?.state === "pending") {
      await pending.promise;
    }

    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Truncated")).toBe(true);
  });

  it("falls back to decoding via <img> when createImageBitmap(blob) throws InvalidStateError", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const bitmap = { width: 200, height: 100 } as any;

    const createImageBitmapSpy = vi.fn((src: any) => {
      if (src instanceof Blob) {
        const err = new Error("decode failed");
        (err as any).name = "InvalidStateError";
        return Promise.reject(err);
      }
      return Promise.resolve(bitmap);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    // jsdom does not implement URL.createObjectURL; stub it while keeping the URL constructor.
    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      // Provide a deterministic Image stub that fires `onload` asynchronously after handlers are installed.
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        width = 16;
        height = 8;
        naturalWidth = 16;
        naturalHeight = 8;
        set src(_value: string) {
          this.onload?.();
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      const renderer = new CanvasGridRenderer({
        provider,
        rowCount: 1,
        colCount: 1,
        defaultColWidth: 100,
        defaultRowHeight: 50,
        imageResolver
      });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(100, 50, 1);

      renderer.renderImmediately();

      const pending = (renderer as any).imageBitmapCache.get("img1") as
        | { state: "pending"; promise: Promise<void> }
        | { state: string };
      // In most environments, the decode will still be pending after the first render pass. In
      // synchronous requestAnimationFrame test setups, it may also complete quickly; handle both.
      if (pending?.state === "pending") {
        await pending.promise;
      } else {
        expect(pending?.state).toBe("ready");
      }

      renderer.renderImmediately();

      expect(imageResolver).toHaveBeenCalledTimes(1);
      // One call for blob decode attempt, one call for canvas decode after the fallback.
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
      expect(createImageBitmapSpy.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
      expect(createImageBitmapSpy.mock.calls[1]?.[0]).toBeInstanceOf(HTMLCanvasElement);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
      expect(content.rec.drawImages.length).toBeGreaterThan(0);
    } finally {
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });

  it("falls back to drawing the decoded canvas when createImageBitmap(canvas) throws", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn((src: any) => {
      if (src instanceof Blob) {
        const err = new Error("decode failed");
        (err as any).name = "InvalidStateError";
        return Promise.reject(err);
      }
      if (src instanceof HTMLCanvasElement) {
        return Promise.reject(new Error("bitmap allocation failed"));
      }
      return Promise.resolve({ width: 10, height: 10 } as any);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        width = 16;
        height = 8;
        naturalWidth = 16;
        naturalHeight = 8;
        set src(_value: string) {
          this.onload?.();
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      const renderer = new CanvasGridRenderer({
        provider,
        rowCount: 1,
        colCount: 1,
        defaultColWidth: 100,
        defaultRowHeight: 50,
        imageResolver
      });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(100, 50, 1);

      renderer.renderImmediately();

      const pending = (renderer as any).imageBitmapCache.get("img1") as
        | { state: "pending"; promise: Promise<void> }
        | { state: string };
      if (pending?.state === "pending") {
        await pending.promise;
      } else {
        expect(pending?.state).toBe("ready");
      }

      renderer.renderImmediately();

      expect(imageResolver).toHaveBeenCalledTimes(1);
      // One call for blob decode attempt, one call for the fallback bitmap allocation (which fails).
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
      expect(createImageBitmapSpy.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
      expect(createImageBitmapSpy.mock.calls[1]?.[0]).toBeInstanceOf(HTMLCanvasElement);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
      expect(content.rec.drawImages.length).toBeGreaterThan(0);
      const last = content.rec.drawImages[content.rec.drawImages.length - 1]!;
      expect(last[0]).toBeInstanceOf(HTMLCanvasElement);
    } finally {
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });

  it("revokes object URLs when the <img> fallback fails to load", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const createImageBitmapSpy = vi.fn((src: any) => {
      if (src instanceof Blob) {
        const err = new Error("decode failed");
        (err as any).name = "InvalidStateError";
        return Promise.reject(err);
      }
      return Promise.resolve({ width: 10, height: 10 } as any);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        set src(_value: string) {
          this.onerror?.();
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      const renderer = new CanvasGridRenderer({
        provider,
        rowCount: 1,
        colCount: 1,
        defaultColWidth: 100,
        defaultRowHeight: 50,
        imageResolver
      });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(100, 50, 1);

      renderer.renderImmediately();
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();

      expect(imageResolver).toHaveBeenCalledTimes(1);
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
      expect(createImageBitmapSpy.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    } finally {
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });

  it("revokes object URLs when the <img> fallback times out", async () => {
    vi.useFakeTimers();
    try {
      // Avoid synchronous RAF side effects; this test controls when renders occur.
      vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

      const provider: CellProvider = {
        getCell: (row, col) =>
          row === 0 && col === 0
            ? {
                row,
                col,
                value: null,
                image: { imageId: "img_timeout", altText: "Timeout", width: 100, height: 50 }
              }
            : null
      };

      const invalidState = new Error("decode failed");
      (invalidState as any).name = "InvalidStateError";
      const createImageBitmapSpy = vi.fn(() => Promise.reject(invalidState));
      vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

      // Use a tiny blob (<24 bytes) so the PNG header guard is skipped and we go directly to the InvalidStateError path.
      const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

      const gridCanvas = document.createElement("canvas");
      const contentCanvas = document.createElement("canvas");
      const selectionCanvas = document.createElement("canvas");

      const grid = createRecordingContext(gridCanvas);
      const content = createRecordingContext(contentCanvas);
      const selection = createRecordingContext(selectionCanvas);

      const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
        [gridCanvas, grid.ctx],
        [contentCanvas, content.ctx],
        [selectionCanvas, selection.ctx]
      ]);

      installContexts(contexts);

      const URLCtor = globalThis.URL as any;
      const originalCreateObjectURL = URLCtor?.createObjectURL;
      const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
      const createObjectURL = vi.fn(() => "blob:fake");
      const revokeObjectURL = vi.fn();
      URLCtor.createObjectURL = createObjectURL;
      URLCtor.revokeObjectURL = revokeObjectURL;

      try {
        class FakeImage {
          onload: (() => void) | null = null;
          onerror: (() => void) | null = null;
          set src(_value: string) {
            // Intentionally never call onload/onerror.
          }
        }
        vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

        const renderer = new CanvasGridRenderer({
          provider,
          rowCount: 1,
          colCount: 1,
          defaultColWidth: 100,
          defaultRowHeight: 50,
          imageResolver
        });
        renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
        renderer.resize(100, 50, 1);

        renderer.renderImmediately();

        const pending = (renderer as any).imageBitmapCache.get("img_timeout") as
          | { state: "pending"; promise: Promise<void> }
          | { state: string };
        expect(pending?.state).toBe("pending");

        // Allow createImageBitmap rejection -> fallback decode to schedule its timeout.
        await flushMicrotasks();

        vi.advanceTimersByTime(5_000);
        await flushMicrotasks();

        expect(imageResolver).toHaveBeenCalledTimes(1);
        expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
        expect(createObjectURL).toHaveBeenCalledTimes(1);
        expect(revokeObjectURL).toHaveBeenCalledTimes(1);

        // Final state should be an error (the InvalidStateError is rethrown when the fallback fails).
        expect((renderer as any).imageBitmapCache.get("img_timeout")?.state).toBe("error");

        renderer.renderImmediately();
        expect(content.rec.fillTexts.some((args) => args[0] === "Timeout")).toBe(true);
      } finally {
        if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
        else URLCtor.createObjectURL = originalCreateObjectURL;
        if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
        else URLCtor.revokeObjectURL = originalRevokeObjectURL;
      }
    } finally {
      vi.useRealTimers();
    }
  });

  it("evicts least-recently-used decoded images when the image cache exceeds its max size", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img1", width: 100, height: 50 }
            }
          : row === 0 && col === 1
            ? {
                row,
                col,
                value: null,
                image: { imageId: "img2", altText: "img2", width: 100, height: 50 }
              }
            : null
    };

    const bitmap1 = { width: 200, height: 100, close: vi.fn() } as any;
    const bitmap2 = { width: 200, height: 100, close: vi.fn() } as any;
    const createImageBitmapSpy = vi.fn().mockResolvedValueOnce(bitmap1).mockResolvedValueOnce(bitmap2);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 2,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver,
      imageBitmapCacheMaxEntries: 1
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    // Load + decode the first image.
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);

    // Scroll to the second image so the first is no longer visible.
    renderer.setScroll(100, 0);
    renderer.markAllDirty();
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    // Second decode pushes the cache over the limit; the first bitmap should be evicted+closed.
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
    expect(bitmap1.close).toHaveBeenCalledTimes(1);
    expect(bitmap2.close).not.toHaveBeenCalled();
  });

  it("closes a decoded bitmap when an in-flight request is invalidated before decode completes", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    let resolveDecode!: (value: any) => void;
    const decodePromise = new Promise<any>((resolve) => {
      resolveDecode = resolve;
    });

    const bitmap = { width: 10, height: 10, close: vi.fn() } as any;
    const createImageBitmapSpy = vi.fn(() => decodePromise);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    // Start the request and let it reach the `createImageBitmap` call.
    renderer.renderImmediately();
    await flushMicrotasks();
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);

    // Invalidate before the decode finishes; the decoded bitmap should be closed
    // once it resolves (since nothing can consume it anymore).
    renderer.invalidateImage("img1");

    resolveDecode(bitmap);
    await flushMicrotasks();

    expect(bitmap.close).toHaveBeenCalledTimes(1);
  });

  it("closes a decoded bitmap when clearImageCache is called while a decode is in-flight", async () => {
    // Avoid synchronous RAF side effects; this test controls when renders occur.
    vi.stubGlobal("requestAnimationFrame", (_cb: FrameRequestCallback) => 0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    let resolveDecode!: (value: any) => void;
    const decodePromise = new Promise<any>((resolve) => {
      resolveDecode = resolve;
    });

    const bitmap1 = { width: 10, height: 10, close: vi.fn() } as any;
    const bitmap2 = { width: 10, height: 10, close: vi.fn() } as any;
    const createImageBitmapSpy = vi.fn().mockReturnValueOnce(decodePromise).mockResolvedValueOnce(bitmap2);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    await flushMicrotasks();
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);

    const pending = (renderer as any).imageBitmapCache.get("img1") as { state: "pending"; promise: Promise<void> } | undefined;
    expect(pending?.state).toBe("pending");

    renderer.clearImageCache();

    resolveDecode(bitmap1);
    await (pending?.promise ?? Promise.resolve());
    await flushMicrotasks();

    expect(bitmap1.close).toHaveBeenCalledTimes(1);

    // Re-render should start a new decode (bitmap2) and keep it cached.
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
    expect(bitmap2.close).not.toHaveBeenCalled();
  });

  it("dedupes image requests and marks content dirty when decoding completes", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const bitmap = { width: 10, height: 10 } as any;
    const createImageBitmapSpy = vi.fn(async () => bitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    let resolveSource: ((value: Blob | null) => void) | undefined;
    const imageResolver = vi.fn(
      () =>
        new Promise<Blob | null>((resolve) => {
          resolveSource = resolve;
        })
    );

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 2,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(200, 50, 1);

    const markDirtySpy = vi.spyOn((renderer as any).dirty.content, "markDirty");
    markDirtySpy.mockClear();

    renderer.renderImmediately();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(0);
    expect(markDirtySpy).toHaveBeenCalledTimes(0);

    if (!resolveSource) {
      throw new Error("Expected image resolver promise to be captured.");
    }
    resolveSource(new Blob([new Uint8Array([1])], { type: "image/png" }));
    await flushMicrotasks();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
    expect(markDirtySpy).toHaveBeenCalled();
  });

  it("retries after a decode failure (without permanently poisoning the cache)", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const bitmap = { width: 10, height: 10 } as any;
    const createImageBitmapSpy = vi
      .fn()
      .mockRejectedValueOnce(new Error("decode failed"))
      .mockResolvedValueOnce(bitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    // Initial attempt: decode fails and is cached briefly.
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);

    // Within the negative-cache window, we should *not* retry.
    renderer.renderImmediately();
    await flushMicrotasks();
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);

    // After the retry window, a new render should attempt decoding again.
    vi.advanceTimersByTime(1_000);
    // Force a repaint so the renderer re-requests the image (it only triggers
    // resolution attempts while painting visible cells).
    renderer.markAllDirty();
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
    expect(content.rec.drawImages.length).toBeGreaterThan(0);

    vi.useRealTimers();
  });

  it("invalidateImage clears the error retry window so images can be retried immediately", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    try {
      const provider: CellProvider = {
        getCell: (row, col) =>
          row === 0 && col === 0
            ? {
                row,
                col,
                value: null,
                image: { imageId: "img1", altText: "img", width: 100, height: 50 }
              }
            : null
      };
 
      const bitmap = { width: 10, height: 10 } as any;
      const createImageBitmapSpy = vi
        .fn()
        .mockRejectedValueOnce(new Error("decode failed"))
        .mockResolvedValueOnce(bitmap);
      vi.stubGlobal("createImageBitmap", createImageBitmapSpy);
 
      const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));
 
      const gridCanvas = document.createElement("canvas");
      const contentCanvas = document.createElement("canvas");
      const selectionCanvas = document.createElement("canvas");
 
      const grid = createRecordingContext(gridCanvas);
      const content = createRecordingContext(contentCanvas);
      const selection = createRecordingContext(selectionCanvas);
 
      const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
        [gridCanvas, grid.ctx],
        [contentCanvas, content.ctx],
        [selectionCanvas, selection.ctx]
      ]);
 
      HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
        const existing = contexts.get(this);
        if (existing) return existing;
        const created = createRecordingContext(this).ctx;
        contexts.set(this, created);
        return created;
      }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
 
      const renderer = new CanvasGridRenderer({
        provider,
        rowCount: 1,
        colCount: 1,
        defaultColWidth: 100,
        defaultRowHeight: 50,
        imageResolver
      });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(100, 50, 1);
 
      // Initial attempt: decode fails and is cached briefly.
      renderer.renderImmediately();
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();
 
      expect(imageResolver).toHaveBeenCalledTimes(1);
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
 
      // Calling invalidateImage should drop the error entry and allow an immediate retry,
      // even if we have not advanced time past the retry window.
      renderer.invalidateImage("img1");
 
      // Flush the new async request and repaint.
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();
 
      expect(imageResolver).toHaveBeenCalledTimes(2);
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
      expect(content.rec.drawImages.length).toBeGreaterThan(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("retries after an imageResolver failure (without permanently poisoning the cache)", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);
    try {
      const provider: CellProvider = {
        getCell: (row, col) =>
          row === 0 && col === 0
            ? {
                row,
                col,
                value: null,
                image: { imageId: "img1", altText: "img", width: 100, height: 50 }
              }
            : null
      };

      const bitmap = { width: 10, height: 10 } as any;
      const createImageBitmapSpy = vi.fn().mockResolvedValue(bitmap);
      vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

      const imageResolver = vi
        .fn()
        .mockRejectedValueOnce(new Error("resolver failed"))
        .mockResolvedValueOnce(new Blob([new Uint8Array([1])], { type: "image/png" }));

      const gridCanvas = document.createElement("canvas");
      const contentCanvas = document.createElement("canvas");
      const selectionCanvas = document.createElement("canvas");

      const grid = createRecordingContext(gridCanvas);
      const content = createRecordingContext(contentCanvas);
      const selection = createRecordingContext(selectionCanvas);

      const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
        [gridCanvas, grid.ctx],
        [contentCanvas, content.ctx],
        [selectionCanvas, selection.ctx]
      ]);

      HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
        const existing = contexts.get(this);
        if (existing) return existing;
        const created = createRecordingContext(this).ctx;
        contexts.set(this, created);
        return created;
      }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

      const renderer = new CanvasGridRenderer({
        provider,
        rowCount: 1,
        colCount: 1,
        defaultColWidth: 100,
        defaultRowHeight: 50,
        imageResolver
      });
      renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
      renderer.resize(100, 50, 1);

      // Initial attempt: the resolver fails and is cached briefly.
      renderer.renderImmediately();
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();

      expect(imageResolver).toHaveBeenCalledTimes(1);
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(0);

      // Within the negative-cache window, we should *not* retry the resolver.
      renderer.renderImmediately();
      await flushMicrotasks();
      expect(imageResolver).toHaveBeenCalledTimes(1);

      // After the retry window, a new render should attempt resolution again.
      vi.advanceTimersByTime(1_000);
      renderer.markAllDirty();
      renderer.renderImmediately();
      await flushMicrotasks();
      renderer.renderImmediately();
      await flushMicrotasks();

      expect(imageResolver).toHaveBeenCalledTimes(2);
      expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
      expect(content.rec.drawImages.length).toBeGreaterThan(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("clearImageCache closes ready bitmaps and forces a re-decode", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    const bitmap1 = { width: 10, height: 10, close: vi.fn() } as any;
    const bitmap2 = { width: 10, height: 10, close: vi.fn() } as any;
    const createImageBitmapSpy = vi.fn().mockResolvedValueOnce(bitmap1).mockResolvedValueOnce(bitmap2);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    installContexts(contexts);

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50,
      imageResolver
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    // Decode + draw the initial bitmap.
    renderer.renderImmediately();
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
    expect(bitmap1.close).not.toHaveBeenCalled();

    // Clearing should close the ready bitmap and allow a new decode.
    renderer.clearImageCache();
    expect(bitmap1.close).toHaveBeenCalledTimes(1);

    // Flush the new decode triggered by the repaint.
    await flushMicrotasks();
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(2);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(2);
    expect(bitmap2.close).not.toHaveBeenCalled();
  });

  it("clearImageCache does not loop indefinitely when requestAnimationFrame is synchronous", async () => {
    const provider: CellProvider = {
      getCell: (row, col) =>
        row === 0 && col === 0
          ? {
              row,
              col,
              value: null,
              image: { imageId: "img1", altText: "img", width: 100, height: 50 }
            }
          : null
    };

    // Guard against regressions where clearImageCache triggers an infinite loop of
    // invalidate->requestRender->renderFrame->requestImageBitmap->reinsert while iterating.
    const rafSpy = vi.fn((cb: FrameRequestCallback) => {
      if (rafSpy.mock.calls.length > 25) {
        throw new Error("requestAnimationFrame called too many times (possible clearImageCache loop)");
      }
      cb(0);
      return 0;
    });
    vi.stubGlobal("requestAnimationFrame", rafSpy);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const grid = createRecordingContext(gridCanvas);
    const content = createRecordingContext(contentCanvas);
    const selection = createRecordingContext(selectionCanvas);

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>([
      [gridCanvas, grid.ctx],
      [contentCanvas, content.ctx],
      [selectionCanvas, selection.ctx]
    ]);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      const existing = contexts.get(this);
      if (existing) return existing;
      const created = createRecordingContext(this).ctx;
      contexts.set(this, created);
      return created;
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const renderer = new CanvasGridRenderer({
      provider,
      rowCount: 1,
      colCount: 1,
      defaultColWidth: 100,
      defaultRowHeight: 50
    });
    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.resize(100, 50, 1);

    renderer.renderImmediately();
    await flushMicrotasks();

    // Sanity: the renderer should have created an entry for the visible image.
    expect(((renderer as any).imageBitmapCache as Map<string, unknown>).size).toBeGreaterThan(0);

    rafSpy.mockClear();
    renderer.clearImageCache();

    // Ensure we did not spin the RAF loop excessively.
    expect(rafSpy.mock.calls.length).toBeLessThan(25);
  });
});
