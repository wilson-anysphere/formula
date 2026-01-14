import { afterEach, describe, expect, it, vi } from "vitest";

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
    // Placeholder rendering uses these Canvas APIs.
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
      pos: { xEmu: pxToEmu(4), yEmu: pxToEmu(6) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

async function flushMicrotasks(times = 4): Promise<void> {
  for (let idx = 0; idx < times; idx++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("DrawingOverlay images", () => {
  it("dedupes in-flight decodes across render passes and requests a redraw when ready", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const entry: ImageEntry = { id: "img_1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const images: ImageStore = {
      get: (id) => (id === entry.id ? entry : undefined),
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    let resolveDecode!: (bitmap: ImageBitmap) => void;
    const decodePromise = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });
    const createImageBitmapMock = vi.fn(() => decodePromise);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const geom: GridGeometry = {
      cellOriginPx: () => ({ x: 0, y: 0 }),
      cellSizePx: () => ({ width: 0, height: 0 }),
    };

    const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

    const requestRender = vi.fn();
    const overlay = new DrawingOverlay(canvas, images, geom, undefined, requestRender);

    const objects = [createImageObject(entry.id)];

    overlay.render(objects, viewport);
    overlay.render(objects, viewport);

    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(0);
    expect(requestRender).toHaveBeenCalledTimes(0);

    resolveDecode({} as ImageBitmap);

    await flushMicrotasks();

    expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
    expect(requestRender).toHaveBeenCalledTimes(1);

    overlay.render(objects, viewport);
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(1);
  });
});
