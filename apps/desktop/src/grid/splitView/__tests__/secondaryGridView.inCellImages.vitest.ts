/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import type { ImageStore } from "../../../drawings/types";
import { SecondaryGridView } from "../secondaryGridView";

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

describe("SecondaryGridView in-cell images", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    vi.stubGlobal("createImageBitmap", vi.fn(async () => ({ width: 1, height: 1 } as any)));
  });

  const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

  it("forwards the image resolver to the underlying CanvasGridRenderer", async () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 120 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 60 });
    document.body.appendChild(container);

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // Doc coords (0-based). Grid coords will be +1/+1 for headers.
    doc.setCellValue(sheetId, { row: 0, col: 0 }, { type: "image", value: { imageId: " img1 ", altText: " Img " } });

    const imageResolver = vi.fn(async () => new Blob([new Uint8Array([1])], { type: "image/png" }));

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 6,
      colCount: 6,
      showFormulas: () => false,
      getComputedValue: () => null,
      imageResolver,
      getDrawingObjects: () => [],
      images,
    });

    gridView.grid.renderer.renderImmediately();
    await flushMicrotasks();
    gridView.grid.renderer.renderImmediately();

    expect(imageResolver).toHaveBeenCalledWith("img1");

    gridView.destroy();
    container.remove();
  });
});
