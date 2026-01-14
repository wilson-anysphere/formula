/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu } from "../../drawings/overlay";
import { buildHitTestIndex, drawingObjectToViewportRect, hitTestDrawings } from "../../drawings/hitTest";
import { getResizeHandleCenters } from "../../drawings/selectionHandles";
import type { DrawingObject, ImageStore } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorCanvasCharts: string | undefined;
let priorUseCanvasCharts: string | undefined;

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
      rotate: vi.fn(),
      translate: vi.fn(),
      scale: vi.fn(),
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

describe("SpreadsheetApp drawing overlay (legacy grid)", () => {
  let priorCanvasCharts: string | undefined;
  let priorUseCanvasCharts: string | undefined;

  afterEach(() => {
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;
  });

  beforeEach(() => {
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.USE_CANVAS_CHARTS = "0";
    document.body.innerHTML = "";

    // This suite covers legacy chart selection overlay behavior; disable canvas charts so
    // the dedicated selection overlay canvas is mounted.
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.USE_CANVAS_CHARTS = "0";

    // Avoid leaking `?canvasCharts=...` URL params between test suites.
    try {
      const url = new URL(window.location.href);
      url.search = "";
      url.hash = "";
      window.history.replaceState(null, "", url.toString());
    } catch {
      // ignore history errors
    }

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // jsdom (as used by vitest) does not provide PointerEvent in all environments.
    // SpreadsheetApp only relies on MouseEvent fields (clientX/Y, button) for drawing hit tests.
    if (typeof (globalThis as any).PointerEvent === "undefined") {
      Object.defineProperty(globalThis, "PointerEvent", { configurable: true, value: MouseEvent });
    }

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

  it("mounts between the base + content canvases and resizes with DPR", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const resizeSpy = vi.spyOn(DrawingOverlay.prototype, "resize");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("legacy");

      const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      expect(drawingCanvas).not.toBeNull();

      const gridCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--base");
      const contentCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--content");
      expect(gridCanvas).not.toBeNull();
      expect(contentCanvas).not.toBeNull();

      const children = Array.from(root.children);
      const gridIdx = children.indexOf(gridCanvas!);
      const drawingIdx = children.indexOf(drawingCanvas!);
      const contentIdx = children.indexOf(contentCanvas!);
      expect(gridIdx).toBeGreaterThanOrEqual(0);
      expect(drawingIdx).toBeGreaterThan(gridIdx);
      expect(drawingIdx).toBeLessThan(contentIdx);

      expect(resizeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          width: 800,
          height: 600,
          dpr: 2,
        }),
      );
      expect(drawingCanvas!.width).toBe(800 * 2);
      expect(drawingCanvas!.height).toBe(600 * 2);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("passes frozen pane counts from the document even if derived frozen counts are stale", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const renderSpy = vi.spyOn((app as any).drawingOverlay, "render");
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      // Simulate stale cached counts (charts historically had a bug here).
      (app as any).frozenRows = 0;
      (app as any).frozenCols = 0;
      (app as any).scrollX = 50;
      (app as any).scrollY = 100;

      renderSpy.mockClear();
      app.refresh("scroll");

      expect(renderSpy).toHaveBeenCalled();
      const lastCall = renderSpy.mock.calls.at(-1);
      const viewport = lastCall?.[1] as any;

      expect(viewport).toEqual(
        expect.objectContaining({
          frozenRows: 1,
          frozenCols: 1,
        }),
      );

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not throw when DrawingOverlay.render throws", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
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
    process.env.DESKTOP_GRID_MODE = "legacy";
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
    process.env.DESKTOP_GRID_MODE = "legacy";
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
    process.env.DESKTOP_GRID_MODE = "legacy";
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

  it("computes consistent render vs interaction viewports for drawings (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
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
      app.setScroll(50, 10);

      const renderViewport = app.getDrawingRenderViewport();
      const interactionViewport = app.getDrawingInteractionViewport();

      // Render and interaction viewports share the same coordinate space (full grid-root).
      expect(renderViewport.headerOffsetX).toBe(interactionViewport.headerOffsetX);
      expect(renderViewport.headerOffsetY).toBe(interactionViewport.headerOffsetY);
      expect(renderViewport.frozenWidthPx).toBe(interactionViewport.frozenWidthPx);
      expect(renderViewport.frozenHeightPx).toBe(interactionViewport.frozenHeightPx);

      // Verify hit testing aligns with where the object is rendered in drawingCanvas space.
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

  it("passes frozen pane metadata to the drawing overlay viewport (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const renderSpy = vi.spyOn((app as any).drawingOverlay, "render");
      expect(app.getGridMode()).toBe("legacy");

      const doc = app.getDocument();
      renderSpy.mockClear();
      doc.setFrozen(app.getCurrentSheetId(), 1, 2, { label: "Freeze" });

      expect(renderSpy).toHaveBeenCalled();
      const viewport = renderSpy.mock.calls.at(-1)?.[1] as any;

      expect(viewport).toEqual(
        expect.objectContaining({
          frozenRows: 1,
          frozenCols: 2,
        }),
      );

      // Frozen extents are expressed in full root coordinates (including the row/col headers),
      // matching the drawing canvas coordinate space.
      const firstScrollableCol = app.getCellRectA1("C1");
      const firstScrollableRow = app.getCellRectA1("A2");
      expect(firstScrollableCol).not.toBeNull();
      expect(firstScrollableRow).not.toBeNull();

      expect(viewport.frozenWidthPx).toBe(firstScrollableCol!.x);
      expect(viewport.frozenHeightPx).toBe(firstScrollableRow!.y);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("selects drawings on pointerdown using the interaction viewport (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
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

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      // Row/col headers are 48px/24px in SpreadsheetApp; click inside the drawing just under them.
      const event = new PointerEvent("pointerdown", {
        bubbles: true,
        cancelable: true,
        clientX: 48 + 5,
        clientY: 24 + 5,
        button: 0,
      });
      selectionCanvas!.dispatchEvent(event);

      expect(app.getSelectedDrawingId()).toBe(1);
      // In legacy mode, drawing selection chrome is rendered on the selection canvas; the overlay stays unselected.
      expect(((app as any).drawingOverlay as any).selectedId).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("re-renders the selection canvas when selecting a drawing via selectDrawingById (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });

      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
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

      const renderSelectionSpy = vi.spyOn(app as any, "renderSelection");
      renderSelectionSpy.mockClear();

      app.selectDrawingById(1);

      expect(app.getSelectedDrawingId()).toBe(1);
      // In legacy mode, selection chrome is drawn on the selection canvas, so the drawing overlay stays unselected.
      expect(((app as any).drawingOverlay as any).selectedId).toBe(null);
      expect(renderSelectionSpy).toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("renders rotated drawing selections in the selection canvas layer", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const rotated: DrawingObject = {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
        transform: { rotationDeg: 90, flipH: false, flipV: false },
      };

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => [rotated];

      (app as any).selectedDrawingId = 1;
      const selectionCtx = (app as any).selectionCtx as any;
      selectionCtx.rotate.mockClear();

      (app as any).renderSelection();

      expect(selectionCtx.rotate).toHaveBeenCalledWith(Math.PI / 2);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("starts a resize gesture when clicking a rotated resize handle", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);

      const rotated: DrawingObject = {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
        transform: { rotationDeg: 90, flipH: false, flipV: false },
      };

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => [rotated];

      const viewport = (app as any).getDrawingInteractionViewport();
      const geom = (app as any).drawingGeom;
      const bounds = drawingObjectToViewportRect(rotated, viewport, geom);
      const handleCenter = getResizeHandleCenters(bounds, rotated.transform).find((c) => c.handle === "nw");
      expect(handleCenter).toBeTruthy();

      (app as any).onPointerDown({
        clientX: handleCenter!.x,
        clientY: handleCenter!.y,
        pointerType: "mouse",
        button: 0,
        pointerId: 1,
        ctrlKey: false,
        metaKey: false,
        shiftKey: false,
        altKey: false,
        preventDefault: () => {},
      } as any);

      expect((app as any).drawingGesture).toEqual(expect.objectContaining({ mode: "resize", handle: "nw" }));

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("locks image aspect ratio when Shift is held during corner resize (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      // Avoid exercising bitmap decode paths in this unit test; we only care about the
      // interaction math + cached anchor updates.
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
      const handleCenter = getResizeHandleCenters(bounds, image.transform).find((c) => c.handle === "se");
      expect(handleCenter).toBeTruthy();

      (app as any).onPointerDown({
        clientX: handleCenter!.x,
        clientY: handleCenter!.y,
        pointerType: "mouse",
        button: 0,
        pointerId: 1,
        ctrlKey: false,
        metaKey: false,
        shiftKey: false,
        altKey: false,
        preventDefault: () => {},
      } as any);

      // Drag horizontally while holding Shift; height should adjust to keep 2:1 ratio.
      (app as any).onPointerMove({
        clientX: handleCenter!.x + 50,
        clientY: handleCenter!.y,
        pointerId: 1,
        shiftKey: true,
      } as any);

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
});
