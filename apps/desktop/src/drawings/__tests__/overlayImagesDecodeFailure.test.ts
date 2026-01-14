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
    setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
    strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
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

function createImageObject(id: number, imageId: string, x: number, y: number): DrawingObject {
  return {
    id,
    kind: { type: "image", imageId },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(x), yEmu: pxToEmu(y) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: id,
  };
}

function createImageStore(entries: ImageEntry[]): ImageStore {
  const map = new Map<string, ImageEntry>(entries.map((entry) => [entry.id, entry]));
  return {
    get: (id) => map.get(id),
    set: (entry) => map.set(entry.id, entry),
  };
}

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("DrawingOverlay image decode failures", () => {
  it("renders placeholders instead of rejecting, and retries decoding after failures", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images = createImageStore([
      { id: "bad", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/bad" },
      { id: "good", bytes: new Uint8Array([4, 5, 6]), mimeType: "image/good" },
    ]);

    let failBad = true;
    const createImageBitmapMock = vi.fn((blob: Blob) => {
      if (blob.type === "image/bad" && failBad) {
        return Promise.reject(new Error("decode failed"));
      }
      return Promise.resolve({} as ImageBitmap);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const overlay = new DrawingOverlay(canvas, images, geom);
    const objects = [createImageObject(0, "bad", 5, 7), createImageObject(1, "good", 30, 40)];

    await expect(overlay.render(objects, viewport)).resolves.toBeUndefined();
    // The good image should still render even if another image fails to decode.
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(1);
    // The bad image should render as a placeholder instead of rejecting the render pass.
    expect(calls.some((c) => c.method === "strokeRect")).toBe(true);
    expect(calls.some((c) => c.method === "fillText" && c.args[0] === "image")).toBe(true);

    // Let the negative-cache entry expire, then allow the decode to succeed and ensure we retry.
    vi.advanceTimersByTime(300);
    vi.setSystemTime(300);
    failBad = false;
    calls.length = 0;

    await expect(overlay.render(objects, viewport)).resolves.toBeUndefined();
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(3);
    expect(createImageBitmapMock.mock.calls.filter(([blob]) => (blob as Blob).type === "image/bad")).toHaveLength(2);
  });

  it("handles synchronous createImageBitmap throws the same way as promise rejections", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const images = createImageStore([
      { id: "bad", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/bad" },
      { id: "good", bytes: new Uint8Array([4, 5, 6]), mimeType: "image/good" },
    ]);

    let failBad = true;
    const createImageBitmapMock = vi.fn((blob: Blob) => {
      if (blob.type === "image/bad" && failBad) {
        throw new Error("decode threw");
      }
      return Promise.resolve({} as ImageBitmap);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const overlay = new DrawingOverlay(canvas, images, geom);
    const objects = [createImageObject(0, "bad", 5, 7), createImageObject(1, "good", 30, 40)];

    await expect(overlay.render(objects, viewport)).resolves.toBeUndefined();
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(1);
    expect(calls.some((c) => c.method === "strokeRect")).toBe(true);
    expect(calls.some((c) => c.method === "fillText" && c.args[0] === "image")).toBe(true);

    // Expire the negative-cache entry and ensure we retry decoding.
    vi.advanceTimersByTime(300);
    vi.setSystemTime(300);
    failBad = false;
    calls.length = 0;

    await expect(overlay.render(objects, viewport)).resolves.toBeUndefined();
    expect(calls.filter((c) => c.method === "drawImage")).toHaveLength(2);
    expect(createImageBitmapMock).toHaveBeenCalledTimes(3);
    expect(createImageBitmapMock.mock.calls.filter(([blob]) => (blob as Blob).type === "image/bad")).toHaveLength(2);
  });
});
