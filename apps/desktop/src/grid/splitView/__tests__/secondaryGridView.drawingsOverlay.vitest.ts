/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../secondaryGridView";
import { DocumentController } from "../../../document/documentController.js";
import type { DrawingObject, ImageStore } from "../../../drawings/types";
import { ImageBitmapCache } from "../../../drawings/imageBitmapCache";
import { DrawingOverlay, pxToEmu } from "../../../drawings/overlay";

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

function createRecordingCanvasContext(calls: string[]): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const ctx: any = {
    canvas: document.createElement("canvas"),
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,

    // DrawingOverlay calls
    clearRect: noop,
    save: noop,
    restore: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    setLineDash: noop,
    setTransform: noop,
    drawImage: noop,
    strokeRect: () => calls.push("strokeRect"),
    fillText: () => calls.push("fillText"),
  };

  return new Proxy(ctx, {
    get(target, prop) {
      if (prop in target) return (target as any)[prop];
      return noop;
    },
    set(target, prop, value) {
      (target as any)[prop] = value;
      return true;
    },
  }) as any;
}

describe("SecondaryGridView drawings overlay", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0,
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("renders drawing objects to the drawings canvas layer", async () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 300 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 200 });
    document.body.appendChild(container);

    const calls: string[] = [];
    const drawingsCtx = createRecordingCanvasContext(calls);

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function (this: HTMLCanvasElement) {
        if (this.classList.contains("grid-canvas--drawings")) return drawingsCtx;
        return createMockCanvasContext();
      },
    });

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const images: ImageStore = { get: () => undefined, set: () => {} };

    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "shape" },
        zOrder: 0,
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(40), cy: pxToEmu(20) },
        },
      },
    ];

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 20,
      colCount: 20,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => objects,
      images,
    });

    await (gridView as any).renderDrawings();

    expect(calls).toContain("strokeRect");
    expect(calls).toContain("fillText");

    gridView.destroy();
    container.remove();
  });

  it("passes frozen pane metadata to the drawings overlay viewport", async () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 300 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 200 });
    document.body.appendChild(container);

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    // Freeze 1 row + 2 cols (sheet-level counts; excludes the shared-grid headers).
    doc.setFrozen(sheetId, 1, 2, { label: "Freeze" });

    const images: ImageStore = { get: () => undefined, set: () => {} };
    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 20,
      colCount: 20,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");
    renderSpy.mockClear();

    await (gridView as any).renderDrawings();

    expect(renderSpy).toHaveBeenCalled();
    const viewport = renderSpy.mock.calls.at(-1)?.[1] as any;

    const renderer = gridView.grid.renderer;
    const gridViewport = renderer.scroll.getViewportState();
    const headerOffsetX = renderer.scroll.cols.totalSize(1);
    const headerOffsetY = renderer.scroll.rows.totalSize(1);

    expect(viewport).toEqual(
      expect.objectContaining({
        frozenRows: 1,
        frozenCols: 2,
        headerOffsetX,
        headerOffsetY,
        frozenWidthPx: gridViewport.frozenWidth,
        frozenHeightPx: gridViewport.frozenHeight,
      }),
    );

    gridView.destroy();
    container.remove();
  });

  it("releases drawing overlay bitmap caches on destroy()", () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 300 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 200 });
    document.body.appendChild(container);

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    const clearSpy = vi.spyOn(ImageBitmapCache.prototype, "clear");

    const doc = new DocumentController();
    const images: ImageStore = { get: () => undefined, set: () => {} };

    const view = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => "Sheet1",
      rowCount: 20,
      colCount: 20,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => [],
      images,
    });

    view.destroy();
    expect(clearSpy).toHaveBeenCalled();

    container.remove();
  });
});
