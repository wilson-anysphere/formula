import { afterEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { ImageBitmapCache } from "../imageBitmapCache";
import type { DrawingObject, ImageEntry, ImageStore } from "../types";

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function createStubCanvasContext(): {
  ctx: CanvasRenderingContext2D;
  calls: Array<{ method: string; args: unknown[] }>;
} {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const ctx: any = {
    clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
    drawImage: (...args: unknown[]) => calls.push({ method: "drawImage", args }),
    save: () => calls.push({ method: "save", args: [] }),
    restore: () => calls.push({ method: "restore", args: [] }),
    beginPath: () => calls.push({ method: "beginPath", args: [] }),
    rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
    clip: () => calls.push({ method: "clip", args: [] }),
    setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
    strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
    fill: () => calls.push({ method: "fill", args: [] }),
    stroke: () => calls.push({ method: "stroke", args: [] }),
  };

  return { ctx: ctx as CanvasRenderingContext2D, calls };
}

function createStubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  const canvas: any = {
    width: 0,
    height: 0,
    style: {},
    getContext: (type: string) => (type === "2d" ? ctx : null),
  };
  return canvas as HTMLCanvasElement;
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

const imageEntry: ImageEntry = {
  id: "img_1",
  bytes: new Uint8Array([1, 2, 3]),
  mimeType: "image/png",
};

const images: ImageStore = {
  get: (id) => (id === imageEntry.id ? imageEntry : undefined),
  set: () => {},
};

function createImageObject({ x, y, id = 1, zOrder = 0 }: { x: number; y: number; id?: number; zOrder?: number }): DrawingObject {
  return {
    id,
    kind: { type: "image", imageId: imageEntry.id },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(x), yEmu: pxToEmu(y) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder,
  };
}

function createShapeObject({ x, y, id = 2, zOrder = 1 }: { x: number; y: number; id?: number; zOrder?: number }): DrawingObject {
  return {
    id,
    kind: { type: "shape" },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(x), yEmu: pxToEmu(y) },
      size: { cx: pxToEmu(10), cy: pxToEmu(10) },
    },
    zOrder,
  };
}

describe("DrawingOverlay async render races", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("does not paint stale image/selection overlay when earlier decode finishes after a later render", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const overlay = new DrawingOverlay(createStubCanvas(ctx), images, geom);
    overlay.setSelectedId(1);

    const first = deferred<ImageBitmap>();
    const second = deferred<ImageBitmap>();

    vi.spyOn(ImageBitmapCache.prototype, "get")
      .mockImplementationOnce(() => first.promise)
      .mockImplementationOnce(() => second.promise);

    const p1 = overlay.render([createImageObject({ x: 5, y: 6 })], viewport);
    const p2 = overlay.render([createImageObject({ x: 25, y: 30 })], viewport);

    const bitmap = {} as ImageBitmap;
    second.resolve(bitmap);
    await p2;

    first.resolve(bitmap);
    await p1;

    const drawCalls = calls.filter((call) => call.method === "drawImage");
    expect(drawCalls).toHaveLength(1);
    expect(drawCalls[0]?.args).toEqual([bitmap, 25, 30, 20, 10]);

    const selectionCalls = calls.filter((call) => call.method === "strokeRect");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toEqual([25, 30, 20, 10]);
  });

  it("does not paint stale placeholder overlays when earlier decode finishes after a later render", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const overlay = new DrawingOverlay(createStubCanvas(ctx), images, geom);

    const first = deferred<ImageBitmap>();
    const second = deferred<ImageBitmap>();

    vi.spyOn(ImageBitmapCache.prototype, "get")
      .mockImplementationOnce(() => first.promise)
      .mockImplementationOnce(() => second.promise);

    const p1 = overlay.render(
      [createImageObject({ x: 0, y: 0, zOrder: 0 }), createShapeObject({ x: 10, y: 12, zOrder: 1 })],
      viewport,
    );
    const p2 = overlay.render(
      [createImageObject({ x: 0, y: 0, zOrder: 0 }), createShapeObject({ x: 40, y: 41, zOrder: 1 })],
      viewport,
    );

    const bitmap = {} as ImageBitmap;
    second.resolve(bitmap);
    await p2;

    first.resolve(bitmap);
    await p1;

    const placeholderCalls = calls.filter((call) => call.method === "strokeRect");
    expect(placeholderCalls).toHaveLength(1);
    expect(placeholderCalls[0]?.args).toEqual([40, 41, 10, 10]);
  });
});

