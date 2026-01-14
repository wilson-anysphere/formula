/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { drawingObjectToViewportRect } from "../../drawings/hitTest";
import { pxToEmu } from "../../drawings/overlay";
import { getRotationHandleCenter } from "../../drawings/selectionHandles";
import type { DrawingObject } from "../../drawings/types";
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

describe("SpreadsheetApp drawing hover cursor", () => {
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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("shows a resize cursor when hovering a drawing corner handle", () => {
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
      expect(app.getGridMode()).toBe("shared");

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      };
      (app as any).drawingObjects = [drawing];

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const headerOffsetX = (app as any).rowHeaderWidth ?? 48;
      const headerOffsetY = (app as any).colHeaderHeight ?? 24;
      const x = headerOffsetX + 100;
      const y = headerOffsetY + 80;
      (app as any).onSharedPointerMove({
        clientX: x,
        clientY: y,
        offsetX: x,
        offsetY: y,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);

      expect(root.style.cursor).toBe("nwse-resize");
      expect(selectionCanvas.style.cursor).toBe("nwse-resize");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("shows a grab cursor when hovering the selected drawing rotation handle", () => {
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
      expect(app.getGridMode()).toBe("shared");

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      };

      // Ensure `renderDrawings` can pick up the seeded drawing objects.
      const docAny = (app as any).document as any;
      docAny.getSheetDrawings = () => [drawing];

      (app as any).selectedDrawingId = drawing.id;
      (app as any).renderDrawings();

      const viewport = app.getDrawingInteractionViewport();
      const bounds = drawingObjectToViewportRect(drawing, viewport, (app as any).drawingGeom);
      const handleCenter = getRotationHandleCenter(bounds, drawing.transform);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      (app as any).onSharedPointerMove({
        clientX: handleCenter.x,
        clientY: handleCenter.y,
        offsetX: handleCenter.x,
        offsetY: handleCenter.y,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);

      expect(root.style.cursor).toBe("grab");
      expect(selectionCanvas.style.cursor).toBe("grab");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not show a grab cursor for a scrollable drawing when hovering inside a frozen pane", () => {
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
      expect(app.getGridMode()).toBe("shared");

      const sheetId = app.getCurrentSheetId();
      const doc: any = app.getDocument();
      // Freeze the first column so the sheet has a frozen pane region under the header.
      doc.setFrozen(sheetId, 0, 1, { label: "Freeze First Column" });
      (app as any).syncFrozenPanes();

      const viewportBeforeScroll = app.getDrawingInteractionViewport();
      expect(viewportBeforeScroll.frozenCols).toBe(1);
      expect(viewportBeforeScroll.frozenWidthPx ?? 0).toBeGreaterThan(viewportBeforeScroll.headerOffsetX ?? 0);

      // Position the drawing so its left edge is exactly at the scrollable pane boundary.
      // Then scroll right by enough pixels that the (unclipped) rotation handle center would land
      // inside the frozen pane region.
      const headerOffsetX = Number.isFinite(viewportBeforeScroll.headerOffsetX) ? Math.max(0, viewportBeforeScroll.headerOffsetX!) : 0;
      const frozenBoundaryX = Number.isFinite(viewportBeforeScroll.frozenWidthPx)
        ? Math.max(headerOffsetX, Math.min(viewportBeforeScroll.frozenWidthPx!, viewportBeforeScroll.width))
        : headerOffsetX;
      const frozenCellWidth = frozenBoundaryX - headerOffsetX;

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(frozenCellWidth), yEmu: pxToEmu(200) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      };
      app.setDrawingObjects([drawing]);
      app.selectDrawing(drawing.id);

      app.setScroll(30, 0);

      const viewport = app.getDrawingInteractionViewport();
      const bounds = drawingObjectToViewportRect(drawing, viewport, (app as any).drawingGeom);
      const handleCenter = getRotationHandleCenter(bounds, drawing.transform);

      // Sanity check: the handle center is in the frozen pane region (under the row header),
      // but the drawing is scrollable (absolute anchors are always treated as scrollable).
      const frozenBoundaryAfterScroll = Number.isFinite(viewport.frozenWidthPx)
        ? Math.max(headerOffsetX, Math.min(viewport.frozenWidthPx!, viewport.width))
        : headerOffsetX;
      expect(handleCenter.x).toBeGreaterThan(headerOffsetX);
      expect(handleCenter.x).toBeLessThan(frozenBoundaryAfterScroll);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const priorCanvasCursor = selectionCanvas.style.cursor;
      root.style.cursor = "crosshair";

      (app as any).onSharedPointerMove({
        clientX: handleCenter.x,
        clientY: handleCenter.y,
        offsetX: handleCenter.x,
        offsetY: handleCenter.y,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);

      expect(root.style.cursor).toBe("");
      expect(selectionCanvas.style.cursor).toBe(priorCanvasCursor);
      expect(selectionCanvas.style.cursor).not.toBe("grab");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not show a rotation cursor for chart drawings (rotation handle disabled)", () => {
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
      expect(app.getGridMode()).toBe("shared");

      const chartDrawing: DrawingObject = {
        id: 1,
        kind: { type: "chart", chartId: "chart_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      };

      // Seed the drawing list and select it so `drawingCursorAtPoint` checks selection handles.
      app.setDrawingObjects([chartDrawing]);
      app.selectDrawing(chartDrawing.id);

      const viewport = app.getDrawingInteractionViewport();
      const bounds = drawingObjectToViewportRect(chartDrawing, viewport, (app as any).drawingGeom);
      const handleCenter = getRotationHandleCenter(bounds, chartDrawing.transform);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const priorCanvasCursor = selectionCanvas.style.cursor;
      root.style.cursor = "crosshair";

      (app as any).onSharedPointerMove({
        clientX: handleCenter.x,
        clientY: handleCenter.y,
        offsetX: handleCenter.x,
        offsetY: handleCenter.y,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);

      expect(root.style.cursor).toBe("");
      expect(selectionCanvas.style.cursor).toBe(priorCanvasCursor);
      expect(selectionCanvas.style.cursor).not.toBe("grab");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("does not show a resize cursor when hovering a clipped selection handle in another frozen pane", () => {
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
      expect(app.getGridMode()).toBe("shared");

      // Freeze the first column so the selection handles for scrollable objects are clipped.
      const docAny = (app as any).document as any;
      docAny.setFrozen?.((app as any).sheetId, 0, 1, { label: "Freeze first column" });
      (app as any).syncFrozenPanes?.();

      expect(app.getFrozen().frozenCols).toBe(1);

      const viewport = app.getDrawingInteractionViewport();
      const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
      const frozenBoundaryX = Number.isFinite(viewport.frozenWidthPx) ? (viewport.frozenWidthPx as number) : headerOffsetX;
      expect(frozenBoundaryX).toBeGreaterThan(headerOffsetX);
      // Place a drawing so it spans the frozen boundary, but the top-left handle lies inside the frozen pane.
      const posX = viewport.scrollX + (frozenBoundaryX - headerOffsetX) - 10;

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img_1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(posX), yEmu: pxToEmu(80) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      };
      app.setDrawingObjects([drawing]);
      app.selectDrawing(drawing.id);
      (app as any).renderDrawings();

      const bounds = drawingObjectToViewportRect(drawing, app.getDrawingInteractionViewport(), (app as any).drawingGeom);
      expect(bounds.x).toBeGreaterThanOrEqual(headerOffsetX);
      expect(bounds.x).toBeLessThan(frozenBoundaryX);
      expect(bounds.x + bounds.width).toBeGreaterThan(frozenBoundaryX);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const priorCanvasCursor = selectionCanvas.style.cursor;
      root.style.cursor = "crosshair";

      // Hover the top-left handle center, which is in the frozen pane and clipped away by the overlay.
      (app as any).onSharedPointerMove({
        clientX: bounds.x,
        clientY: bounds.y,
        offsetX: bounds.x,
        offsetY: bounds.y,
        buttons: 0,
        pointerType: "mouse",
        target: selectionCanvas,
      } as any);

      expect(root.style.cursor).toBe("");
      expect(selectionCanvas.style.cursor).toBe(priorCanvasCursor);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
