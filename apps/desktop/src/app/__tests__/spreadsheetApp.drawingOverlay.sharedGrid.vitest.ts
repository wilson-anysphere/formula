/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu } from "../../drawings/overlay";
import { buildHitTestIndex, hitTestDrawings } from "../../drawings/hitTest";
import type { DrawingObject, ImageStore } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

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

function createRoot(): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

describe("SpreadsheetApp drawing overlay (shared grid)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 2 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("mounts the drawing canvas, resizes with DPR, and re-renders on shared-grid scroll + sheet change", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const resizeSpy = vi.spyOn(DrawingOverlay.prototype, "resize");
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const canvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      expect(canvas).not.toBeNull();

      expect(resizeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          width: 800 - 48,
          height: 600 - 24,
          dpr: 2,
        }),
      );

      // `DrawingOverlay.resize` should size the backing buffer at DPR scale.
      expect(canvas!.width).toBe((800 - 48) * 2);
      expect(canvas!.height).toBe((600 - 24) * 2);

      renderSpy.mockClear();
      const sharedGrid = (app as any).sharedGrid;
      sharedGrid.scrollTo(0, 100);
      expect(renderSpy).toHaveBeenCalled();

      // Switching sheets should trigger a drawings overlay rerender so objects can refresh.
      const doc = app.getDocument();
      doc.addSheet({ sheetId: "sheet_2", name: "Sheet2" });
      doc.setCellValue("sheet_2", { row: 0, col: 0 }, "Seed2");
      renderSpy.mockClear();
      app.activateSheet("sheet_2");
      expect(renderSpy).toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("passes frozen pane metadata to the drawing overlay viewport so drawings pin + clip correctly", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      renderSpy.mockClear();
      const sharedGrid = (app as any).sharedGrid;
      sharedGrid.scrollTo(50, 100);

      expect(renderSpy).toHaveBeenCalled();
      const gridViewport = sharedGrid.renderer.scroll.getViewportState();
      const headerWidth = sharedGrid.renderer.scroll.cols.totalSize(1);
      const headerHeight = sharedGrid.renderer.scroll.rows.totalSize(1);
      const offsetX = Math.min(headerWidth, gridViewport.width);
      const offsetY = Math.min(headerHeight, gridViewport.height);
      const cellAreaWidth = Math.max(0, gridViewport.width - offsetX);
      const cellAreaHeight = Math.max(0, gridViewport.height - offsetY);

      // `DrawingOverlay.prototype.render` is shared between the sheet drawings overlay and the
      // chart selection overlay. Find the call corresponding to the drawings canvas viewport,
      // whose origin is positioned under the headers (so headerOffsetX/Y are 0).
      const viewport = renderSpy.mock.calls
        .map((call) => call?.[1] as any)
        .find((vp) => vp && vp.headerOffsetX === 0 && vp.headerOffsetY === 0 && vp.width === cellAreaWidth && vp.height === cellAreaHeight);
      expect(viewport).toBeTruthy();

      expect(viewport).toEqual(
        expect.objectContaining({
          frozenRows: 1,
          frozenCols: 1,
        }),
      );

      const expectedFrozenWidthPx = Math.min(cellAreaWidth, Math.max(0, gridViewport.frozenWidth - offsetX));
      const expectedFrozenHeightPx = Math.min(cellAreaHeight, Math.max(0, gridViewport.frozenHeight - offsetY));

      expect(viewport.frozenWidthPx).toBe(expectedFrozenWidthPx);
      expect(viewport.frozenHeightPx).toBe(expectedFrozenHeightPx);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("renders per-sheet drawings + images from DocumentController", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockResolvedValue(undefined);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      // Inject a simple workbook drawing layer: one floating image anchored to A1.
      const imageId = "image1.png";
      const bytes = new Uint8Array([1, 2, 3]);
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setImage(imageId, { bytes, mimeType: "image/png" });
      doc.setSheetDrawings(sheetId, [
        {
          id: "d1",
          zOrder: 0,
          anchor: { type: "cell", sheetId, row: 0, col: 0 },
          kind: { type: "image", imageId },
          size: { width: 120, height: 80 },
        },
      ]);

      renderSpy.mockClear();

      // Force a drawing render pass and assert that the overlay receives our object.
      (app as any).renderDrawings();
      expect(renderSpy).toHaveBeenCalled();
      const objects = renderSpy.mock.calls[0]?.[0] as any[];
      expect(objects).toHaveLength(1);
      expect(objects[0]).toMatchObject({ kind: { type: "image", imageId } });

      // Ensure the overlay image store is backed by the document's image map.
      const imageStore = (app as any).drawingImages;
      expect(imageStore.get(imageId)).toMatchObject({ id: imageId, mimeType: "image/png" });
      expect(imageStore.get(imageId)?.bytes).toEqual(bytes);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("computes consistent render vs interaction viewports for drawings (headers + frozen panes)", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const sharedGrid = (app as any).sharedGrid;

      // Scroll so the object is in the scrollable quadrant and uses scroll offsets.
      sharedGrid.scrollTo(50, 10);
      const sharedViewport = sharedGrid.renderer.scroll.getViewportState();

      const renderViewport = app.getDrawingRenderViewport(sharedViewport);
      const interactionViewport = app.getDrawingInteractionViewport(sharedViewport);

      // Render viewport (drawingCanvas-local).
      expect(renderViewport.headerOffsetX).toBe(0);
      expect(renderViewport.headerOffsetY).toBe(0);
      // Interaction viewport (selectionCanvas/root-local).
      expect(interactionViewport.headerOffsetX).toBeGreaterThan(0);
      expect(interactionViewport.headerOffsetY).toBeGreaterThan(0);

      // Frozen boundaries should map between viewport spaces by the header offsets.
      expect(interactionViewport.frozenWidthPx! - interactionViewport.headerOffsetX!).toBe(renderViewport.frozenWidthPx);
      expect(interactionViewport.frozenHeightPx! - interactionViewport.headerOffsetY!).toBe(renderViewport.frozenHeightPx);

      // Verify hit testing aligns with the rectangle rendered in drawingCanvas space.
      const geom = (app as any).drawingOverlay.geom;
      const images: ImageStore = { get: () => undefined, set: () => {} };

      const object: DrawingObject = {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 2, col: 2 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(20) },
        },
        zOrder: 0,
      };

      const calls: Array<{ method: string; args: unknown[] }> = [];
      const ctx: any = {
        clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
        save: () => calls.push({ method: "save", args: [] }),
        restore: () => calls.push({ method: "restore", args: [] }),
        beginPath: () => calls.push({ method: "beginPath", args: [] }),
        rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
        clip: () => calls.push({ method: "clip", args: [] }),
        setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
        strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
        fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
      };

      const canvas: any = {
        width: 0,
        height: 0,
        style: {},
        getContext: () => ctx,
      };

      const overlay = new DrawingOverlay(canvas as HTMLCanvasElement, images, geom);
      await overlay.render([object], renderViewport);

      const stroke = calls.find((c) => c.method === "strokeRect");
      expect(stroke).toBeTruthy();
      const [renderX, renderY, w, h] = stroke!.args as number[];

      const index = buildHitTestIndex([object], geom, { bucketSizePx: 64 });
      const hit = hitTestDrawings(
        index,
        interactionViewport,
        renderX + interactionViewport.headerOffsetX! + 1,
        renderY + interactionViewport.headerOffsetY! + 1,
      );
      expect(hit?.object.id).toBe(1);
      expect(hit?.bounds).toEqual({
        x: renderX + interactionViewport.headerOffsetX!,
        y: renderY + interactionViewport.headerOffsetY!,
        width: w,
        height: h,
      });

      // The hit-test bounds should map back to the render-space bounds by subtracting headers.
      expect(hit!.bounds.x - interactionViewport.headerOffsetX!).toBe(renderX);
      expect(hit!.bounds.y - interactionViewport.headerOffsetY!).toBe(renderY);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
