import { describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageEntry, ImageStore } from "../types";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: Array<{ method: string; args: unknown[] }> } {
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

function createImageObject(opts: { id: number; imageId: string; zOrder: number; x: number; y: number }): DrawingObject {
  return {
    id: opts.id,
    kind: { type: "image", imageId: opts.imageId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(opts.x), yEmu: pxToEmu(opts.y) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: opts.zOrder,
  };
}

function createImageStore(entries: Record<string, ImageEntry>): ImageStore {
  return {
    get: (id) => entries[id],
    set: () => {},
  };
}

function createTrackedThenable<T>(opts: {
  onThen: () => void;
}): { thenable: Promise<T>; resolve: (value: T) => void; reject: (err: unknown) => void } {
  let resolve!: (value: T) => void;
  let reject!: (err: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });

  const thenable: any = {
    then: (onFulfilled: (value: T) => unknown, onRejected: (err: unknown) => unknown) => {
      opts.onThen();
      return promise.then(onFulfilled, onRejected);
    },
    // `DrawingOverlay` may attach a `catch` handler to prefetched decode promises to avoid
    // unhandled rejections. Provide it here so our lightweight thenable behaves like a Promise.
    catch: (onRejected: (err: unknown) => unknown) => promise.catch(onRejected),
  };

  return { thenable: thenable as Promise<T>, resolve, reject };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

describe("DrawingOverlay images", () => {
  it("prefetches visible image bitmaps concurrently and draws in z-order", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images = createImageStore({
      img1: { id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
      img2: { id: "img2", bytes: new Uint8Array([4, 5, 6]), mimeType: "image/png" },
    });

    const overlay = new DrawingOverlay(canvas, images, geom);

    let getCalls = 0;
    const getCallsAtAwait: number[] = [];
    const controlsById = new Map<string, ReturnType<typeof createTrackedThenable<ImageBitmap>>>();

    const getMock = vi.fn((entry: ImageEntry) => {
      getCalls += 1;
      const control = createTrackedThenable<ImageBitmap>({
        onThen: () => {
          getCallsAtAwait.push(getCalls);
        },
      });
      controlsById.set(entry.id, control);
      return control.thenable;
    });

    (overlay as any).bitmapCache = { get: getMock };

    const obj1 = createImageObject({ id: 1, imageId: "img1", zOrder: 0, x: 5, y: 7 });
    const obj2 = createImageObject({ id: 2, imageId: "img2", zOrder: 1, x: 15, y: 17 });

    const bitmap1 = { tag: "bitmap1" } as unknown as ImageBitmap;
    const bitmap2 = { tag: "bitmap2" } as unknown as ImageBitmap;

    const renderPromise = overlay.render([obj1, obj2], viewport);

    expect(getMock).toHaveBeenCalledTimes(2);
    // `await`ing a thenable triggers `.then(...)` via a microtask; yield once so we
    // can observe the first awaited decode.
    await Promise.resolve();
    expect(getCallsAtAwait).toEqual([2]);

    // Resolve out-of-order to ensure draw order respects zOrder.
    controlsById.get("img2")!.resolve(bitmap2);
    expect(calls.some((call) => call.method === "drawImage")).toBe(false);

    controlsById.get("img1")!.resolve(bitmap1);
    await renderPromise;

    const drawCalls = calls.filter((call) => call.method === "drawImage");
    expect(drawCalls).toHaveLength(2);
    expect(drawCalls[0]!.args[0]).toBe(bitmap1);
    expect(drawCalls[1]!.args[0]).toBe(bitmap2);
  });

  it("aborts stale renders after awaiting image decode", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images = createImageStore({
      img1: { id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
    });

    const overlay = new DrawingOverlay(canvas, images, geom);

    const control = createTrackedThenable<ImageBitmap>({ onThen: () => {} });
    const getMock = vi.fn(() => control.thenable);
    (overlay as any).bitmapCache = { get: getMock };

    const obj = createImageObject({ id: 1, imageId: "img1", zOrder: 0, x: 5, y: 7 });

    const firstRender = overlay.render([obj], viewport);
    await overlay.render([], viewport);

    control.resolve({ tag: "bitmap1" } as unknown as ImageBitmap);
    await firstRender;

    expect(calls.some((call) => call.method === "drawImage")).toBe(false);
  });

  it("falls back to placeholder when image decode fails", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images = createImageStore({
      img1: { id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
    });

    const overlay = new DrawingOverlay(canvas, images, geom);

    const getMock = vi.fn(() => {
      const rejected = Promise.reject(new Error("boom"));
      // Attach a handler immediately to avoid Node's unhandled rejection warnings;
      // the overlay will still `await` the original rejection and render a placeholder.
      rejected.catch(() => {});
      return rejected;
    });
    (overlay as any).bitmapCache = { get: getMock };

    const obj = createImageObject({ id: 1, imageId: "img1", zOrder: 0, x: 5, y: 7 });
    await overlay.render([obj], viewport);

    expect(calls.some((call) => call.method === "drawImage")).toBe(false);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(true);
  });
});
