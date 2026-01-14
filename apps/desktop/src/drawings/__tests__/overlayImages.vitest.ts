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
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("swallows async requestRender rejections (prevents unhandled promise rejections)", async () => {
    const { ctx } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = {
      id: "img_async_requestRender",
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
    vi.stubGlobal("createImageBitmap", vi.fn(() => bitmapPromise));

    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown) => {
      unhandled.push(reason);
    };
    process.on("unhandledRejection", onUnhandled);

    const requestRender = vi.fn(async () => {
      throw new Error("boom");
    });
    const overlay = new DrawingOverlay(canvas, images, geom, undefined, requestRender);

    try {
      overlay.render([createImageObject(entry.id)], viewport);

      resolveBitmap({ close: () => {} } as unknown as ImageBitmap);
      await flushMicrotasks();

      // Allow Node a turn to emit any unhandled rejection events.
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(requestRender).toHaveBeenCalledTimes(1);
      expect(unhandled).toHaveLength(0);
    } finally {
      process.off("unhandledRejection", onUnhandled);
    }
  });

  it("swallows async render rejections when image hydration triggers a rerender without requestRender", async () => {
    const { ctx } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = {
      id: "img_async_render",
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
    vi.stubGlobal("createImageBitmap", vi.fn(() => bitmapPromise));

    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown) => {
      unhandled.push(reason);
    };
    process.on("unhandledRejection", onUnhandled);

    // No requestRender callback => scheduleHydrationRerender falls back to calling `render()` itself.
    const overlay = new DrawingOverlay(canvas, images, geom);

    try {
      overlay.render([createImageObject(entry.id)], viewport);

      // Simulate a unit test mocking `render()` as async after the initial paint has captured the
      // last render args (so the hydration rerender will call our stub).
      const asyncRender = vi.fn(async () => {
        throw new Error("boom");
      });
      (overlay as any).render = asyncRender;

      resolveBitmap({ close: () => {} } as unknown as ImageBitmap);
      await flushMicrotasks();
      // Allow Node a turn to emit any unhandled rejection events.
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(asyncRender).toHaveBeenCalledTimes(1);
      expect(unhandled).toHaveLength(0);
    } finally {
      process.off("unhandledRejection", onUnhandled);
    }
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

  it("does not attempt bitmap decode for zero-size drawing anchors", () => {
    const { ctx } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = {
      id: "img_zero",
      bytes: new Uint8Array([1, 2, 3]),
      mimeType: "image/png",
    };
    const images: ImageStore = {
      get: (id) => (id === entry.id ? entry : undefined),
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    const createImageBitmapSpy = vi.fn(async () => ({ close: () => {} }) as unknown as ImageBitmap);
    vi.stubGlobal("createImageBitmap", createImageBitmapSpy);

    const requestRender = vi.fn();
    const overlay = new DrawingOverlay(canvas, images, geom, undefined, requestRender);

    const obj: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: entry.id },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };

    overlay.render([obj], viewport);

    expect(createImageBitmapSpy).not.toHaveBeenCalled();
    expect(requestRender).not.toHaveBeenCalled();
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
