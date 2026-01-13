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

function createImageObject(id: number, imageId: string): DrawingObject {
  return {
    id,
    kind: { type: "image", imageId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
      size: { cx: pxToEmu(10), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

describe("DrawingOverlay render ordering", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("does not draw results from an earlier async render after a newer render completes", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imagesById = new Map<string, ImageEntry>([
      ["a", { id: "a", bytes: new Uint8Array([1]), mimeType: "image/a" }],
      ["b", { id: "b", bytes: new Uint8Array([2]), mimeType: "image/b" }],
    ]);

    const images: ImageStore = {
      get: (id) => imagesById.get(id),
      set: (entry) => {
        imagesById.set(entry.id, entry);
      },
    };

    const bitmapA = { tag: "A" } as unknown as ImageBitmap;
    const bitmapB = { tag: "B" } as unknown as ImageBitmap;

    let resolveA!: (value: ImageBitmap) => void;
    const deferredA = new Promise<ImageBitmap>((res) => {
      resolveA = res;
    });

    const createImageBitmapMock = vi.fn((blob: Blob) => {
      if (blob.type === "image/a") return deferredA;
      if (blob.type === "image/b") return Promise.resolve(bitmapB);
      return Promise.reject(new Error(`Unexpected image mime type: ${blob.type}`));
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const overlay = new DrawingOverlay(canvas, images, geom);

    const renderA = overlay.render([createImageObject(1, "a")], viewport);
    await overlay.render([createImageObject(2, "b")], viewport);

    resolveA(bitmapA);
    await renderA;

    const drawCalls = calls.filter((call) => call.method === "drawImage");
    expect(drawCalls.length).toBeGreaterThan(0);

    const drawBitmaps = drawCalls.map((call) => call.args[0]);
    const lastDraw = drawBitmaps[drawBitmaps.length - 1];
    expect(lastDraw).toBe(bitmapB);

    const lastBIndex = drawBitmaps.lastIndexOf(bitmapB);
    expect(lastBIndex).toBe(drawBitmaps.length - 1);
    expect(drawBitmaps.slice(lastBIndex + 1)).not.toContain(bitmapA);
  });
});

