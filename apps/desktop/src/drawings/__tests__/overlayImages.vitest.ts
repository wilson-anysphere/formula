import { afterEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageEntry, ImageStore } from "../types";

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
    fillRect: (...args: unknown[]) => calls.push({ method: "fillRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
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

function createImageObject(imageId: string): DrawingObject {
  return {
    id: 1,
    kind: { type: "image", imageId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

async function flushMicrotasks(times = 4): Promise<void> {
  for (let idx = 0; idx < times; idx++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("DrawingOverlay images", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("renders synchronously and draws the image after decode + requestRender", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = {
      id: "img_1",
      bytes: new Uint8Array([1, 2, 3]),
      mimeType: "image/png",
    };
    const images: ImageStore = {
      get: (id) => (id === entry.id ? entry : undefined),
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    let resolveBitmap!: (bitmap: ImageBitmap) => void;
    const bitmapPromise = new Promise<ImageBitmap>((resolve) => {
      resolveBitmap = resolve;
    });
    const createImageBitmapSpy = vi.fn(() => bitmapPromise);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const requestRender = vi.fn();
    const overlay = new DrawingOverlay(canvas, images, geom, undefined, requestRender);

    const result = overlay.render([createImageObject(entry.id)], viewport);
    expect(result).toBeUndefined();

    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
    expect(calls.some((call) => call.method === "drawImage")).toBe(false);
    expect(requestRender).not.toHaveBeenCalled();

    const fakeBitmap = { close: () => {} } as unknown as ImageBitmap;
    resolveBitmap(fakeBitmap);
    await flushMicrotasks();

    expect(requestRender).toHaveBeenCalledTimes(1);
    expect(calls.some((call) => call.method === "drawImage")).toBe(false);

    overlay.render([createImageObject(entry.id)], viewport);
    expect(calls.some((call) => call.method === "drawImage")).toBe(true);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
  });

  it("caches decode failures to avoid infinite retries", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = {
      id: "img_fail",
      bytes: new Uint8Array([9, 9, 9]),
      mimeType: "image/png",
    };
    const images: ImageStore = {
      get: (id) => (id === entry.id ? entry : undefined),
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    const createImageBitmapSpy = vi.fn(async () => {
      throw new Error("decode failed");
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const requestRender = vi.fn();
    const overlay = new DrawingOverlay(canvas, images, geom, undefined, requestRender);

    // First render should not throw, even though the decode will fail.
    expect(() => overlay.render([createImageObject(entry.id)], viewport)).not.toThrow();
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
    expect(calls.some((call) => call.method === "drawImage")).toBe(false);

    await flushMicrotasks();

    // Subsequent renders should not retry decoding.
    overlay.render([createImageObject(entry.id)], viewport);
    expect(createImageBitmapSpy).toHaveBeenCalledTimes(1);
    expect(calls.some((call) => call.method === "drawImage")).toBe(false);
  });
});
