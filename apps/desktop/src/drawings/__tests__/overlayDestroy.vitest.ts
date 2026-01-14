import { afterEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu, type ChartRenderer, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageEntry, ImageStore } from "../types";

function createStubCanvasContext(): CanvasRenderingContext2D {
  const ctx: any = {
    clearRect: () => {},
    drawImage: () => {},
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    rect: () => {},
    clip: () => {},
    setLineDash: () => {},
    strokeRect: () => {},
    fillText: () => {},
    ellipse: () => {},
    moveTo: () => {},
    lineTo: () => {},
    arcTo: () => {},
    closePath: () => {},
    fill: () => {},
    stroke: () => {},
    measureText: (text: string) => ({ width: text.length * 6 }),
  };
  return ctx as CanvasRenderingContext2D;
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

function createShapeObject(id: number, raw_xml: string): DrawingObject {
  return {
    id,
    kind: { type: "shape", raw_xml },
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

describe("DrawingOverlay destroy()", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("closes cached ImageBitmaps", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn(() => Promise.resolve(bitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imageEntry: ImageEntry = { id: "img_1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const entries = new Map<string, ImageEntry>([[imageEntry.id, imageEntry]]);
    const images: ImageStore = {
      get: (id: string) => entries.get(id),
      set: (entry: ImageEntry) => entries.set(entry.id, entry),
      delete: (id: string) => entries.delete(id),
      clear: () => entries.clear(),
    };

    const overlay = new DrawingOverlay(canvas, images, geom);
    await overlay.render([createImageObject(imageEntry.id)], viewport);

    overlay.destroy();

    expect(close).toHaveBeenCalledTimes(1);
  });

  it("closes in-flight preload ImageBitmaps when destroyed", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;
    let resolveDecode!: (value: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });

    const createImageBitmapMock = vi.fn(() => inflightDecode);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imageEntry: ImageEntry = { id: "img_preload", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const entries = new Map<string, ImageEntry>([[imageEntry.id, imageEntry]]);
    const images: ImageStore = {
      get: (id: string) => entries.get(id),
      set: (entry: ImageEntry) => entries.set(entry.id, entry),
      delete: (id: string) => entries.delete(id),
      clear: () => entries.clear(),
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    // Start a preload decode but destroy the overlay before it resolves. This ensures the bitmap
    // isn't leaked even though the cache entry is cleared while the decode promise is in-flight.
    const preload = overlay.preloadImage(imageEntry).catch(() => {});
    overlay.destroy();

    resolveDecode(bitmap);
    await preload;
    await Promise.resolve();

    expect(close).toHaveBeenCalledTimes(1);
  });

  it("closes in-flight preload ImageBitmaps when the cache is cleared", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;
    let resolveDecode!: (value: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });

    const createImageBitmapMock = vi.fn(() => inflightDecode);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imageEntry: ImageEntry = { id: "img_clear", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const entries = new Map<string, ImageEntry>([[imageEntry.id, imageEntry]]);
    const images: ImageStore = {
      get: (id: string) => entries.get(id),
      set: (entry: ImageEntry) => entries.set(entry.id, entry),
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    const preload = overlay.preloadImage(imageEntry).catch(() => {});
    overlay.clearImageCache();

    resolveDecode(bitmap);
    await preload;
    await Promise.resolve();

    expect(close).toHaveBeenCalledTimes(1);
  });

  it("closes in-flight preload ImageBitmaps when the image is invalidated", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;
    let resolveDecode!: (value: ImageBitmap) => void;
    const inflightDecode = new Promise<ImageBitmap>((resolve) => {
      resolveDecode = resolve;
    });

    const createImageBitmapMock = vi.fn(() => inflightDecode);
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imageEntry: ImageEntry = { id: "img_invalidate", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const entries = new Map<string, ImageEntry>([[imageEntry.id, imageEntry]]);
    const images: ImageStore = {
      get: (id: string) => entries.get(id),
      set: (entry: ImageEntry) => entries.set(entry.id, entry),
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    const preload = overlay.preloadImage(imageEntry).catch(() => {});
    overlay.invalidateImage(imageEntry.id);

    resolveDecode(bitmap);
    await preload;
    await Promise.resolve();

    expect(close).toHaveBeenCalledTimes(1);
  });

  it("rejects preloadImage after destroy without starting a decode", async () => {
    const close = vi.fn();
    const bitmap = { close } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn(() => Promise.resolve(bitmap));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const imageEntry: ImageEntry = { id: "img_destroyed_preload", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    const images: ImageStore = {
      get: () => imageEntry,
      set: () => {},
    };

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.destroy();

    await expect(overlay.preloadImage(imageEntry)).rejects.toMatchObject({ name: "AbortError" });
    expect(createImageBitmapMock).not.toHaveBeenCalled();
  });

  it("does not cache hydrated image bytes after destroy()", async () => {
    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    let resolveHydration!: (value: ImageEntry | undefined) => void;
    const hydrationPromise = new Promise<ImageEntry | undefined>((resolve) => {
      resolveHydration = resolve;
    });

    const set = vi.fn();
    const getAsync = vi.fn(() => hydrationPromise);
    const images: ImageStore = {
      get: () => undefined,
      set,
      getAsync,
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    await overlay.render([createImageObject("img_hydrate")], viewport);
    // `hydrateImage` calls `getAsync` in a microtask.
    await Promise.resolve();
    expect(getAsync).toHaveBeenCalledTimes(1);

    overlay.destroy();

    resolveHydration({ id: "img_hydrate", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
    await Promise.resolve();

    expect(set).not.toHaveBeenCalled();
  });

  it("prunes cached shape text layouts when objects are removed", async () => {
    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images: ImageStore = {
      get: () => undefined,
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p/>
        </xdr:txBody>
      </xdr:sp>
    `;

    await overlay.render([createShapeObject(123, rawXml)], viewport);
    expect((overlay as any).shapeTextCache.size).toBe(1);

    await overlay.render([], viewport);
    expect((overlay as any).shapeTextCache.size).toBe(0);
  });

  it("clears the spatial index on destroy so drawing objects are not retained", async () => {
    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images: ImageStore = {
      get: () => undefined,
      set: () => {},
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p/>
        </xdr:txBody>
      </xdr:sp>
    `;

    await overlay.render([createShapeObject(123, rawXml)], viewport);
    expect((overlay as any).spatialIndex.getObject(123)).not.toBeNull();

    overlay.destroy();
    expect((overlay as any).spatialIndex.getObject(123)).toBeNull();
  });

  it("calls chartRenderer.destroy() when provided", () => {
    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images: ImageStore = {
      get: () => undefined,
      set: () => {},
      delete: () => {},
      clear: () => {},
    };

    const destroy = vi.fn();
    const chartRenderer: ChartRenderer = {
      renderToCanvas: () => {},
      destroy,
    };

    const overlay = new DrawingOverlay(canvas, images, geom, chartRenderer);
    overlay.destroy();

    expect(destroy).toHaveBeenCalledTimes(1);
  });

  it("clears spatial index state on destroy", async () => {
    const ctx = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images: ImageStore = {
      get: () => undefined,
      set: () => {},
    };

    const overlay = new DrawingOverlay(canvas, images, geom);

    await overlay.render([createShapeObject(42, "<xdr:sp/>")], viewport);
    expect((overlay as any).spatialIndex.getObject(42)).not.toBeNull();

    overlay.destroy();
    expect((overlay as any).spatialIndex.getObject(42)).toBeNull();
  });
});
