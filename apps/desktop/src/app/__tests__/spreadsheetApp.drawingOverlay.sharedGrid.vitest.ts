/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu } from "../../drawings/overlay";
import { buildHitTestIndex, drawingObjectToViewportRect, hitTestDrawings } from "../../drawings/hitTest";
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

function dispatchPointerEvent(
  target: EventTarget,
  type: string,
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number; shiftKey?: boolean },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const base = {
    bubbles: true,
    cancelable: true,
    clientX: opts.clientX,
    clientY: opts.clientY,
    pointerId,
    button,
    shiftKey: opts.shiftKey ?? false,
  };
  const event = (() => {
    const PointerEventCtor = (globalThis as any).PointerEvent;
    if (typeof PointerEventCtor === "function") {
      return new PointerEventCtor(type, base);
    }
    return new MouseEvent(type, base);
  })();
  // Ensure pointer-only fields exist even when the environment shims PointerEvent with MouseEvent.
  // `pointerId` is readonly on real PointerEvent implementations, so set it best-effort.
  try {
    (event as any).pointerId = pointerId;
  } catch {
    // ignore
  }
  target.dispatchEvent(event);
}

describe("SpreadsheetApp drawing overlay (shared grid)", () => {
  let priorCanvasCharts: string | undefined;
  let priorUseCanvasCharts: string | undefined;

  afterEach(() => {
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.USE_CANVAS_CHARTS = "0";
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

    // jsdom does not always provide PointerEvent; SpreadsheetApp listens for pointer events
    // in shared-grid mode. For tests, a MouseEvent-compatible shim is sufficient.
    if (typeof (globalThis as any).PointerEvent === "undefined") {
      Object.defineProperty(globalThis, "PointerEvent", { configurable: true, value: MouseEvent });
    }
    if (typeof (window as any).PointerEvent === "undefined") {
      Object.defineProperty(window, "PointerEvent", { configurable: true, value: (globalThis as any).PointerEvent });
    }

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
          width: 800,
          height: 600,
          dpr: 2,
        }),
      );

      // `DrawingOverlay.resize` should size the backing buffer at DPR scale.
      expect(canvas!.width).toBe(800 * 2);
      expect(canvas!.height).toBe(600 * 2);

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

  it("feeds formula-model drawing objects from DocumentController through the model adapter layer", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    let app: SpreadsheetApp | null = null;
    let root: HTMLElement | null = null;
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockImplementation(() => {});

      root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      app = new SpreadsheetApp(root, status);
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;

      doc.setSheetDrawings(sheetId, [
        {
          id: "12",
          zOrder: 0,
          kind: { Image: { image_id: "image1.png" } },
          anchor: {
            TwoCell: {
              from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
              to: { cell: { row: 1, col: 1 }, offset: { x_emu: 0, y_emu: 0 } },
            },
          },
        },
      ]);

      // Ensure the next render pass re-reads the document state (in case a prior render
      // cached an empty list before drawings were populated).
      (app as any).drawingObjectsCache = null;

      renderSpy.mockClear();
      (app as any).renderDrawings();

      expect(renderSpy).toHaveBeenCalled();
      const objects = renderSpy.mock.calls.at(-1)?.[0] as any[];
      const imageObject = objects.find((obj) => obj?.id === 12) ?? null;
      expect(imageObject).not.toBeNull();
      expect(imageObject).toMatchObject({
        id: 12,
        kind: { type: "image", imageId: "image1.png" },
        anchor: { type: "twoCell" },
        zOrder: 0,
      });
    } finally {
      app?.destroy();
      root?.remove();
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("invalidates shared-grid in-cell image bitmaps when workbook image bytes change via imageDeltas", () => {
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
      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();

      const invalidateSpy = vi.spyOn(sharedGrid.renderer, "invalidateImage");

      const doc: any = app.getDocument();
      doc.applyExternalImageCacheDeltas(
        [
          {
            imageId: "img-1",
            before: null,
            after: { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
          },
        ],
        { source: "collab" },
      );

      expect(invalidateSpy).toHaveBeenCalledWith("img-1");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("clears the shared-grid in-cell image bitmap cache on applyState restores", () => {
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
      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();

      const clearSpy = vi.spyOn(sharedGrid.renderer, "clearImageCache");

      const snapshotDoc = new (app.getDocument().constructor as any)();
      // Ensure Sheet1 is materialized to avoid edge cases where applyState emits minimal deltas.
      snapshotDoc.getCell?.("Sheet1", { row: 0, col: 0 });
      const snapshot = snapshotDoc.encodeState();
      app.getDocument().applyState(snapshot);

      expect(clearSpy).toHaveBeenCalled();

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
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      // Spy on the *drawings* overlay instance specifically. SpreadsheetApp also uses
      // `DrawingOverlay` for chart selection handles, which runs on scroll; using a
      // prototype spy would include those calls and make this test order-dependent.
      const renderSpy = vi.spyOn((app as any).drawingOverlay, "render");

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

      const viewport = renderSpy.mock.calls.at(-1)?.[1] as any;
      expect(viewport).toBeTruthy();

      expect(viewport).toEqual(
        expect.objectContaining({
          width: gridViewport.width,
          height: gridViewport.height,
          headerOffsetX: offsetX,
          headerOffsetY: offsetY,
          frozenRows: 1,
          frozenCols: 1,
          zoom: sharedGrid.renderer.getZoom(),
        }),
      );

      const expectedFrozenWidthPx = Math.min(gridViewport.width, Math.max(offsetX, gridViewport.frozenWidth));
      const expectedFrozenHeightPx = Math.min(gridViewport.height, Math.max(offsetY, gridViewport.frozenHeight));

      // DrawingOverlay viewports express frozen boundaries in *viewport coordinates*. Depending on
      // the overlay (drawings vs chart selection), the viewport may include header offsets.
      // Compare the effective frozen content size (boundary - header offsets) so this test is
      // stable even if other overlays render after drawings.
      const headerOffsetXViewport = typeof viewport.headerOffsetX === "number" ? viewport.headerOffsetX : 0;
      const headerOffsetYViewport = typeof viewport.headerOffsetY === "number" ? viewport.headerOffsetY : 0;
      expect((viewport.frozenWidthPx ?? 0) - headerOffsetXViewport).toBe(expectedFrozenWidthPx - offsetX);
      expect((viewport.frozenHeightPx ?? 0) - headerOffsetYViewport).toBe(expectedFrozenHeightPx - offsetY);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("propagates shared-grid zoom to drawing render + interaction viewports", () => {
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

      const drawingsRenderSpy = vi.spyOn((app as any).drawingOverlay, "render");
      drawingsRenderSpy.mockClear();

      // SpreadsheetApp.setZoom delegates to DesktopSharedGrid which triggers an onScroll callback
      // (and thus a drawings rerender) when zoom changes.
      app.setZoom(2);

      expect(drawingsRenderSpy).toHaveBeenCalled();
      const lastViewport = drawingsRenderSpy.mock.calls.at(-1)?.[1] as any;
      expect(lastViewport?.zoom).toBe(2);

      // Verify the public viewport helpers also include zoom (used by hit testing + interactions).
      expect(app.getDrawingRenderViewport().zoom).toBe(2);
      expect(app.getDrawingInteractionViewport().zoom).toBe(2);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("pickDrawingAtClientPoint hit-tests transformed drawings under zoom", () => {
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
      app.setZoom(2);

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();
      selectionCanvas!.getBoundingClientRect = () =>
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

      const object: DrawingObject = {
        id: 123,
        kind: { type: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
          size: { cx: pxToEmu(100), cy: pxToEmu(100) },
        },
        transform: { rotationDeg: 45, flipH: false, flipV: false },
        zOrder: 0,
      };

      // Provide the UI drawing object directly so SpreadsheetApp doesn't need to decode model objects.
      const doc: any = app.getDocument();
      doc.getSheetDrawings = () => [object];
      (app as any).drawingObjectsCache = null;

      const viewport = app.getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const rect = drawingObjectToViewportRect(object, viewport, geom);

      // Choose a point that is inside the rotated rect but outside the untransformed anchor rect.
      const cx = rect.x + rect.width / 2;
      const cy = rect.y + rect.height / 2;
      const clientX = cx;
      const clientY = cy + (rect.height / 2) * 1.3; // outside untransformed (dy > half-height), inside rotated at 45deg.

      expect(clientY).toBeGreaterThan(rect.y + rect.height);
      expect(app.pickDrawingAtClientPoint(clientX, clientY)).toBe(123);

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
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockImplementation(() => {});

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
      const images = objects.filter((obj) => obj?.kind?.type === "image");
      expect(images).toHaveLength(1);
      expect(images[0]).toMatchObject({ kind: { type: "image", imageId } });

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

  it("does not throw when DrawingOverlay.render throws", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const err = new Error("boom");
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockImplementation(() => {});
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      renderSpy.mockClear();
      warnSpy.mockClear();
      renderSpy.mockImplementationOnce(() => {
        throw err;
      });

      (app as any).renderDrawings();

      expect(warnSpy).toHaveBeenCalledWith("Drawing overlay render failed", err);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not emit an unhandled rejection when DrawingOverlay.render returns a rejected promise", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    vi.useRealTimers();
    try {
      const err = new Error("boom");
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render").mockImplementation(() => {});
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const unhandled: unknown[] = [];
      const onUnhandled = (reason: unknown) => {
        unhandled.push(reason);
      };
      process.on("unhandledRejection", onUnhandled);

      try {
        const app = new SpreadsheetApp(root, status);

        renderSpy.mockClear();
        warnSpy.mockClear();
        renderSpy.mockImplementationOnce(() => Promise.reject(err) as any);

        (app as any).renderDrawings();

        // Allow the rejection handler (and Node unhandledRejection bookkeeping) to run.
        await new Promise((resolve) => setTimeout(resolve, 0));

        expect(unhandled).toHaveLength(0);
        expect(warnSpy).toHaveBeenCalledWith("Drawing overlay render failed", err);

        app.destroy();
      } finally {
        process.off("unhandledRejection", onUnhandled);
        root.remove();
      }
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not throw when chart selection overlay render throws", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    const priorCanvasCharts = process.env.CANVAS_CHARTS;
    const priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.DESKTOP_GRID_MODE = "shared";
    process.env.CANVAS_CHARTS = "0";
    delete process.env.USE_CANVAS_CHARTS;
    try {
      const err = new Error("boom");
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      warnSpy.mockClear();
      // Override just the chart selection overlay instance so drawing overlay renders don't
      // consume the mock implementation.
      const overlay = (app as any).chartSelectionOverlay as DrawingOverlay;
      expect(overlay).toBeTruthy();
      (overlay as any).render = vi.fn(() => {
        throw err;
      });

      (app as any).renderChartSelectionOverlay();

      expect(warnSpy).toHaveBeenCalledWith("Chart selection overlay render failed", err);

      app.destroy();
      root.remove();
    } finally {
      if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = priorCanvasCharts;
      if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
      else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not emit an unhandled rejection when chart selection overlay render returns a rejected promise", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    const priorCanvasCharts = process.env.CANVAS_CHARTS;
    const priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.DESKTOP_GRID_MODE = "shared";
    process.env.CANVAS_CHARTS = "0";
    delete process.env.USE_CANVAS_CHARTS;
    vi.useRealTimers();
    try {
      const err = new Error("boom");
      const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const unhandled: unknown[] = [];
      const onUnhandled = (reason: unknown) => {
        unhandled.push(reason);
      };
      process.on("unhandledRejection", onUnhandled);

      try {
        const app = new SpreadsheetApp(root, status);

        warnSpy.mockClear();
        // Override just the chart selection overlay instance so drawing overlay renders don't
        // consume the mock implementation.
        const overlay = (app as any).chartSelectionOverlay as DrawingOverlay;
        expect(overlay).toBeTruthy();
        (overlay as any).render = vi.fn(() => Promise.reject(err));

        (app as any).renderChartSelectionOverlay();

        // Allow the rejection handler (and Node unhandledRejection bookkeeping) to run.
        await new Promise((resolve) => setTimeout(resolve, 0));

        expect(unhandled).toHaveLength(0);
        expect(warnSpy).toHaveBeenCalledWith("Chart selection overlay render failed", err);

        app.destroy();
      } finally {
        process.off("unhandledRejection", onUnhandled);
        root.remove();
      }
    } finally {
      if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = priorCanvasCharts;
      if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
      else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("computes consistent render vs interaction viewports for drawings (headers + frozen panes)", () => {
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

      // Render and interaction viewports share the same coordinate space (full grid-root).
      expect(renderViewport.headerOffsetX).toBe(interactionViewport.headerOffsetX);
      expect(renderViewport.headerOffsetY).toBe(interactionViewport.headerOffsetY);
      expect(renderViewport.frozenWidthPx).toBe(interactionViewport.frozenWidthPx);
      expect(renderViewport.frozenHeightPx).toBe(interactionViewport.frozenHeightPx);

      // Verify hit testing aligns with the rectangle rendered in drawingCanvas space.
      const geom = (app as any).drawingOverlay.geom;
      const images: ImageStore = { get: () => undefined, set: () => {}, delete: () => {}, clear: () => {} };

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
      overlay.render([object], renderViewport);

      const stroke = calls.find((c) => c.method === "strokeRect");
      expect(stroke).toBeTruthy();
      const [renderX, renderY, w, h] = stroke!.args as number[];

      const index = buildHitTestIndex([object], geom, { bucketSizePx: 64 });
      const hit = hitTestDrawings(
        index,
        interactionViewport,
        renderX + 1,
        renderY + 1,
      );
      expect(hit?.object.id).toBe(1);
      expect(hit?.bounds).toEqual({
        x: renderX,
        y: renderY,
        width: w,
        height: h,
      });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("selects drawings on pointerdown using the interaction viewport (shared grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const selectSpy = vi.spyOn(DrawingOverlay.prototype, "setSelectedId");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      // Add a simple placeholder shape anchored at A1.
      doc.setSheetDrawings(sheetId, [
        {
          id: "1",
          kind: { type: "shape" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: pxToEmu(100), cy: pxToEmu(50) },
          },
          zOrder: 0,
        },
      ]);

      selectSpy.mockClear();

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      // Row/col headers are 48px/24px in SpreadsheetApp; click inside the drawing just under them.
      dispatchPointerEvent(selectionCanvas!, "pointerdown", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 1,
        button: 0,
      });
      dispatchPointerEvent(selectionCanvas!, "pointerup", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 1,
        button: 0,
      });

      expect(selectSpy).toHaveBeenCalledWith(1);

      // Pointerdowns on non-grid overlays (e.g. scrollbars/outline) should not affect drawing selection.
      selectSpy.mockClear();
      const overlay = document.createElement("div");
      root.appendChild(overlay);
      dispatchPointerEvent(overlay, "pointerdown", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 2,
        button: 0,
      });
      dispatchPointerEvent(overlay, "pointerup", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 2,
        button: 0,
      });
      expect(selectSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("ignores non-grid overlay pointerdowns when drawing interactions are enabled (shared grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const selectSpy = vi.spyOn(DrawingOverlay.prototype, "setSelectedId");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      // Add a simple placeholder shape anchored at A1.
      doc.setSheetDrawings(sheetId, [
        {
          id: "1",
          kind: { type: "shape" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
            size: { cx: pxToEmu(100), cy: pxToEmu(50) },
          },
          zOrder: 0,
        },
      ]);

      selectSpy.mockClear();

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      // Row/col headers are 48px/24px in SpreadsheetApp; click inside the drawing just under them.
      dispatchPointerEvent(selectionCanvas!, "pointerdown", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 1,
        button: 0,
      });
      dispatchPointerEvent(selectionCanvas!, "pointerup", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 1,
        button: 0,
      });

      expect(selectSpy).toHaveBeenCalledWith(1);

      // Pointerdowns on non-grid overlays (e.g. scrollbars/outline) should not affect drawing selection
      // even when drawing interactions are enabled.
      selectSpy.mockClear();
      const overlay = document.createElement("div");
      root.appendChild(overlay);
      dispatchPointerEvent(overlay, "pointerdown", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 2,
        button: 0,
      });
      dispatchPointerEvent(overlay, "pointerup", {
        clientX: 48 + 5,
        clientY: 24 + 5,
        pointerId: 2,
        button: 0,
      });
      expect(selectSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("locks image aspect ratio when Shift is held during corner resize (shared grid interactions)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      // Avoid exercising bitmap decode paths in this unit test; we only care about the
      // interaction math + in-memory anchor updates.
      vi.spyOn((app as any).drawingOverlay, "render").mockImplementation(() => {});

      const image: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(200), cy: pxToEmu(100) },
        },
        zOrder: 0,
      };

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => [image];

      const viewport = (app as any).getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const bounds = drawingObjectToViewportRect(image, viewport, geom);

      const seX = bounds.x + bounds.width;
      const seY = bounds.y + bounds.height;

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      // Start resizing from the south-east corner.
      dispatchPointerEvent(selectionCanvas!, "pointerdown", {
        clientX: seX,
        clientY: seY,
        pointerId: 1,
        button: 0,
      });

      // Drag horizontally while holding Shift; height should adjust to keep 2:1 ratio.
      dispatchPointerEvent(selectionCanvas!, "pointermove", {
        clientX: seX + 50,
        clientY: seY,
        pointerId: 1,
        button: 0,
        shiftKey: true,
      });

      const resized = ((app as any).listDrawingObjectsForSheet() as DrawingObject[]).find((obj) => obj.id === 1);
      expect(resized?.anchor).toMatchObject({
        type: "absolute",
        size: { cx: pxToEmu(250), cy: pxToEmu(125) },
      });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("passes shared-grid zoom through to the drawing overlay viewport", () => {
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
      const renderSpy = vi.spyOn((app as any).drawingOverlay, "render");
      expect(app.getZoom()).toBe(1);

      renderSpy.mockClear();
      app.setZoom(2);

      expect(renderSpy).toHaveBeenCalled();
      const viewport = renderSpy.mock.calls.at(-1)?.[1] as any;
      expect(viewport.zoom).toBe(2);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
