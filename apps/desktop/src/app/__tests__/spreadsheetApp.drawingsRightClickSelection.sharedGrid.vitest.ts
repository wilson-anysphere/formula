/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { drawingObjectToViewportRect } from "../../drawings/hitTest";
import { pxToEmu } from "../../drawings/overlay";
import { getRotationHandleCenter } from "../../drawings/selectionHandles";
import type { DrawingObject } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

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
  root.getBoundingClientRect = vi.fn(
    () =>
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
      }) as any,
  );
  document.body.appendChild(root);
  return root;
}

function createPointerLikeMouseEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    button: number;
    pointerId?: number;
    pointerType?: string;
    ctrlKey?: boolean;
    metaKey?: boolean;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button,
    ctrlKey: options.ctrlKey,
    metaKey: options.metaKey,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("SpreadsheetApp drawings right-click selection (shared grid)", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      writable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, writable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("selects the drawing without moving the active cell on right click (without drawing interactions enabled)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    expect(app.getGridMode()).toBe("shared");

    // Move the active cell away from A1 so we can detect selection changes.
    app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
    const beforeActive = app.getActiveCell();

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };

    // Seed draw objects using the public test helper so the capture-based hit test sees them.
    app.setDrawingObjects([drawing]);

    // Right-click within the picture bounds. This should select the drawing but *not* move the active cell.
    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const down = createPointerLikeMouseEvent("pointerdown", { clientX: 60, clientY: 30, button: 2 });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });

  it("selects the drawing without moving the active cell on right click (with drawing interactions enabled)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Right-click selection is handled by the DrawingInteractionController when enabled.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");

    // Move the active cell away from A1 so we can detect selection changes.
    app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
    const beforeActive = app.getActiveCell();

    const sheetId = app.getCurrentSheetId();
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };

    // Use the public test helper to seed draw objects without relying on internal sync helpers.
    app.setDrawingObjects([drawing]);

    // Right-click within the picture bounds. This should select the drawing but *not* move the active
    // cell (Excel-like behavior in shared-grid mode).
    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    // jsdom returns a zero-sized client rect for canvases by default; the drawing interaction controller
    // uses `getBoundingClientRect()` to convert clientX/Y into local coordinates. Ensure the selection
    // canvas reports the same rect as the grid root so hit testing works deterministically.
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const down = createPointerLikeMouseEvent("pointerdown", { clientX: 60, clientY: 30, button: 2 });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });

  it("treats Ctrl+click as a context-click on macOS (drawing interactions enabled)", () => {
    const originalPlatform = navigator.platform;
    const restorePlatform = () => {
      try {
        Object.defineProperty(navigator, "platform", { configurable: true, value: originalPlatform });
      } catch {
        // ignore
      }
    };
    try {
      Object.defineProperty(navigator, "platform", { configurable: true, value: "MacIntel" });
    } catch {
      // If the runtime doesn't allow stubbing `navigator.platform`, skip the test.
      restorePlatform();
      return;
    }

    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      expect(app.getGridMode()).toBe("shared");

      app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
      const beforeActive = app.getActiveCell();

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img-1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
          size: { cx: pxToEmu(100), cy: pxToEmu(100) },
        },
        zOrder: 0,
      };
      app.setDrawingObjects([drawing]);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const bubbled = vi.fn();
      root.addEventListener("pointerdown", bubbled);
      selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

      const down = createPointerLikeMouseEvent("pointerdown", {
        clientX: 60,
        clientY: 30,
        button: 0,
        ctrlKey: true,
        metaKey: false,
      });
      selectionCanvas.dispatchEvent(down);

      expect(app.getSelectedDrawingId()).toBe(1);
      expect(app.getActiveCell()).toEqual(beforeActive);
      expect((down as any).__formulaDrawingContextClick).toBe(true);
      expect(down.defaultPrevented).toBe(false);
      expect(bubbled).toHaveBeenCalledTimes(1);

      app.destroy();
      root.remove();
    } finally {
      restorePlatform();
    }
  });

  it("treats Ctrl+click as a context-click on macOS (drawing interactions disabled)", () => {
    const originalPlatform = navigator.platform;
    const restorePlatform = () => {
      try {
        Object.defineProperty(navigator, "platform", { configurable: true, value: originalPlatform });
      } catch {
        // ignore
      }
    };
    try {
      Object.defineProperty(navigator, "platform", { configurable: true, value: "MacIntel" });
    } catch {
      restorePlatform();
      return;
    }

    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
      expect(app.getGridMode()).toBe("shared");

      app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
      const beforeActive = app.getActiveCell();

      const drawing: DrawingObject = {
        id: 1,
        kind: { type: "image", imageId: "img-1" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(0), yEmu: pxToEmu(0) },
          size: { cx: pxToEmu(100), cy: pxToEmu(100) },
        },
        zOrder: 0,
      };
      app.setDrawingObjects([drawing]);

      const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
      const bubbled = vi.fn();
      root.addEventListener("pointerdown", bubbled);

      const down = createPointerLikeMouseEvent("pointerdown", {
        clientX: 60,
        clientY: 30,
        button: 0,
        ctrlKey: true,
        metaKey: false,
      });
      selectionCanvas.dispatchEvent(down);

      expect(app.getSelectedDrawingId()).toBe(1);
      expect(app.getActiveCell()).toEqual(beforeActive);
      expect((down as any).__formulaDrawingContextClick).toBe(true);
      expect(down.defaultPrevented).toBe(false);
      expect(bubbled).toHaveBeenCalledTimes(1);

      app.destroy();
      root.remove();
    } finally {
      restorePlatform();
    }
  });

  it("keeps selection and tags context-click when right-clicking a selection handle", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // When drawing interactions are disabled, SpreadsheetApp relies on its capture-phase
    // drawing hit testing (`onDrawingPointerDownCapture`) to keep shared-grid selection
    // stable on right-click.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    expect(app.getGridMode()).toBe("shared");

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(1);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Right-click just outside the drawing bounds, but still within the top-left resize handle square.
    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rowHeaderWidth + 100 - 1,
      clientY: colHeaderHeight + 100 - 1,
      button: 2,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });

  it("keeps selection, tags context-click, and focuses the grid when right-clicking a selection handle (drawing interactions enabled)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");

    // Move the active cell away from A1 so we can detect selection changes.
    app.activateCell({ row: 5, col: 5 }, { scrollIntoView: false, focus: false });
    const beforeActive = app.getActiveCell();

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(1);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    // jsdom returns a zero-sized client rect for canvases by default; the drawing interaction controller
    // uses `getBoundingClientRect()` to convert clientX/Y into local coordinates. Ensure the selection
    // canvas reports the same rect as the grid root so hit testing works deterministically.
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);
    const focusSpy = vi.spyOn(root, "focus");

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Right-click just outside the drawing bounds, but still within the top-left resize handle square.
    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: rowHeaderWidth + 100 - 1,
      clientY: colHeaderHeight + 100 - 1,
      button: 2,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);
    expect(focusSpy).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("hitTestDrawingAtClientPoint treats selection handles as drawing hits", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");

    const sheetId = app.getCurrentSheetId();
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };

    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    (app as any).drawingObjectsCache = null;

    // Select the drawing so selection handles are active.
    app.selectDrawingById(1);

    // Top-left resize handle is centered on the top-left corner of the drawing bounds and extends
    // half its size beyond the drawing rect. Right-click slightly outside the rect but within the
    // handle region.
    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;
    const hit = app.hitTestDrawingAtClientPoint(rowHeaderWidth + 100 - 1, colHeaderHeight + 100 - 1);
    expect(hit).toEqual({ id: 1 });

    app.destroy();
    root.remove();
  });

  it("hitTestDrawingAtClientPoint treats the rotation handle as a drawing hit (rotatable drawings)", () => {
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
      kind: { type: "shape", label: "rect" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(1);

    const viewport = app.getDrawingInteractionViewport();
    const bounds = drawingObjectToViewportRect(drawing, viewport, (app as any).drawingGeom);
    const handleCenter = getRotationHandleCenter(bounds, drawing.transform);

    const hit = app.hitTestDrawingAtClientPoint(handleCenter.x, handleCenter.y);
    expect(hit).toEqual({ id: 1 });

    app.destroy();
    root.remove();
  });

  it("hitTestDrawingAtClientPoint does not treat the rotation handle as a hit for chart drawings", () => {
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
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([chartDrawing]);
    app.selectDrawingById(1);

    const viewport = app.getDrawingInteractionViewport();
    const bounds = drawingObjectToViewportRect(chartDrawing, viewport, (app as any).drawingGeom);
    const handleCenter = getRotationHandleCenter(bounds, chartDrawing.transform);

    const hit = app.hitTestDrawingAtClientPoint(handleCenter.x, handleCenter.y);
    expect(hit).toBeNull();

    app.destroy();
    root.remove();
  });

  it("does not treat the rotation handle area as a context-click hit for chart drawings", () => {
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
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([chartDrawing]);
    app.selectDrawingById(1);

    const viewport = app.getDrawingInteractionViewport();
    const bounds = drawingObjectToViewportRect(chartDrawing, viewport, (app as any).drawingGeom);
    const handleCenter = getRotationHandleCenter(bounds, chartDrawing.transform);

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    const down = createPointerLikeMouseEvent("pointerdown", { clientX: handleCenter.x, clientY: handleCenter.y, button: 2 });
    selectionCanvas.dispatchEvent(down);

    // No visible rotation handle for charts, so this should behave like a miss (selection stays, no tag).
    expect(app.getSelectedDrawingId()).toBe(1);
    expect((down as any).__formulaDrawingContextClick).toBeUndefined();
    expect(down.defaultPrevented).toBe(false);

    app.destroy();
    root.remove();
  });
});
