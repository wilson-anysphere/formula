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

async function flushMicrotasks(): Promise<void> {
  // Two turns is usually enough to flush `async` + `finally` chains.
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
    renderer.renderImmediately();
    await flushMicrotasks();

    expect(imageResolver).toHaveBeenCalledTimes(1);
    expect(content.rec.drawImages.length).toBe(0);
    expect(content.rec.fillTexts.some((args) => args[0] === "Missing image")).toBe(true);
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
});
