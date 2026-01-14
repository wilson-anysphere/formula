/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { chartIdToDrawingId } from "../../charts/chartDrawingAdapter";
import { pxToEmu } from "../../drawings/overlay";
import type { DrawingObject } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;
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

describe("SpreadsheetApp canvas charts vs drawing handles (shared grid)", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.DESKTOP_GRID_MODE = "shared";
    process.env.CANVAS_CHARTS = "1";

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
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("does not allow an overlapping chart to steal a context-click on a selected drawing handle", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");
    // This regression relies on canvas charts being handled by a DrawingInteractionController.
    expect((app as any).chartDrawingInteraction).toBeTruthy();

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

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Handle overlap chart",
      position: "A1",
    });
    const chartDrawingId = chartIdToDrawingId(chartId);
    // Place the chart so it overlaps the drawing's bottom-right resize handle.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(160),
      yEmu: pxToEmu(160),
      cxEmu: pxToEmu(140),
      cyEmu: pxToEmu(140),
    });

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    // jsdom returns a zero-sized client rect for canvases by default; interaction controllers
    // use `getBoundingClientRect()` to convert clientX/Y into local coordinates.
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Drawing bottom-right corner is at (100 + 100, 100 + 100) in sheet px. Click just outside
    // the bounds but within the resize handle square so the handle is the intended target.
    const handleClientX = rowHeaderWidth + 200 + 1;
    const handleClientY = colHeaderHeight + 200 + 1;

    // Sanity: without a selected drawing, the chart should be the hit target at this point.
    app.selectDrawingById(null);
    expect(app.hitTestDrawingAtClientPoint(handleClientX, handleClientY)?.id).toBe(chartDrawingId);

    app.selectDrawingById(drawing.id);
    expect(app.getSelectedDrawingId()).toBe(drawing.id);
    expect(app.getSelectedChartId()).toBe(null);
    // With a selected drawing, selection handles should win even though the chart is above.
    expect(app.hitTestDrawingAtClientPoint(handleClientX, handleClientY)?.id).toBe(drawing.id);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    // Right-click on the handle point. This should *not* select the overlapping chart.
    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: handleClientX,
      clientY: handleClientY,
      button: 2,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedDrawingId()).toBe(drawing.id);
    expect(app.getSelectedChartId()).toBe(null);
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });

  it("allows resizing a selected chart via its handles even when a drawing overlaps", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");
    expect((app as any).chartDrawingInteraction).toBeTruthy();

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-under-chart-handle" },
      anchor: {
        type: "absolute",
        // Overlap the chart's bottom-right resize handle area (outside the chart bounds).
        pos: { xEmu: pxToEmu(190), yEmu: pxToEmu(190) },
        size: { cx: pxToEmu(60), cy: pxToEmu(60) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart handle overlap drawing",
      position: "A1",
    });
    const chartDrawingId = chartIdToDrawingId(chartId);
    // Deterministic absolute anchor for stable handle coordinates.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(100),
      yEmu: pxToEmu(100),
      cxEmu: pxToEmu(100),
      cyEmu: pxToEmu(100),
    });

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    // Chart bottom-right corner is at (100 + 100, 100 + 100) in sheet px. Click just outside
    // the bounds but within the resize handle square (and within the overlapping drawing).
    const handleClientX = rowHeaderWidth + 200 + 1;
    const handleClientY = colHeaderHeight + 200 + 1;

    // Sanity: without selecting the chart, this point is a drawing hit (the chart body doesn't cover it).
    expect(app.hitTestDrawingAtClientPoint(handleClientX, handleClientY)?.id).toBe(drawing.id);

    // Select the chart (e.g. via selection pane) so its handles are active.
    app.selectDrawingById(chartDrawingId);
    expect(app.getSelectedChartId()).toBe(chartId);
    expect((app as any).selectedDrawingId).toBeNull();
    // With the chart selected, the resize handle should win even though a drawing overlaps underneath.
    expect(app.hitTestDrawingAtClientPoint(handleClientX, handleClientY)?.id).toBe(chartDrawingId);

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: handleClientX,
      clientY: handleClientY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(down);

    expect(app.getSelectedChartId()).toBe(chartId);
    expect((app as any).selectedDrawingId).toBeNull();
    expect((app as any).chartDrawingGestureActive).toBe(true);
    expect(down.defaultPrevented).toBe(true);

    // End the gesture to avoid leaking pointer listeners across tests.
    const up = createPointerLikeMouseEvent("pointerup", {
      clientX: handleClientX,
      clientY: handleClientY,
      button: 0,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(up);
    expect((app as any).chartDrawingGestureActive).toBe(false);

    app.destroy();
    root.remove();
  });

  it("selects a ChartStore chart over an overlapping drawing on right-click when drawing interactions are disabled", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    expect(app.getGridMode()).toBe("shared");
    // This regression relies on canvas charts being handled by a DrawingInteractionController.
    expect((app as any).chartDrawingInteraction).toBeTruthy();

    // Move the active cell away from A1 so we can detect unwanted selection changes.
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

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Overlapping chart",
      position: "A1",
    });
    const chartDrawingId = chartIdToDrawingId(chartId);
    // Place chart on top of the drawing (canvas charts render above drawings).
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(0),
      yEmu: pxToEmu(0),
      cxEmu: pxToEmu(100),
      cyEmu: pxToEmu(100),
    });

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    // jsdom returns a zero-sized client rect for canvases by default; interaction controllers
    // use `getBoundingClientRect()` to convert clientX/Y into local coordinates.
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;
    const clientX = rowHeaderWidth + 10;
    const clientY = colHeaderHeight + 10;

    // Sanity: hit test should see the chart above the drawing at this point.
    expect(app.hitTestDrawingAtClientPoint(clientX, clientY)?.id).toBe(chartDrawingId);

    const bubbled = vi.fn();
    root.addEventListener("pointerdown", bubbled);

    const down = createPointerLikeMouseEvent("pointerdown", { clientX, clientY, button: 2 });
    selectionCanvas.dispatchEvent(down);

    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(app.getSelectedChartId()).toBe(chartId);
    expect(app.getSelectedDrawingId()).toBe(chartDrawingId);
    expect(app.getActiveCell()).toEqual(beforeActive);
    expect(down.defaultPrevented).toBe(false);
    expect(bubbled).toHaveBeenCalledTimes(1);

    app.destroy();
    root.remove();
  });

  it("preserves ChartStore chart selection on context-click misses", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    expect(app.getGridMode()).toBe("shared");
    expect((app as any).chartDrawingInteraction).toBeTruthy();

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Preserve selection",
      position: "A1",
    });
    // Deterministic absolute anchor for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(0),
      yEmu: pxToEmu(0),
      cxEmu: pxToEmu(100),
      cyEmu: pxToEmu(100),
    });

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    const hitX = rowHeaderWidth + 10;
    const hitY = colHeaderHeight + 10;
    const missX = rowHeaderWidth + 300;
    const missY = colHeaderHeight + 300;

    expect(app.hitTestDrawingAtClientPoint(missX, missY)).toBe(null);

    selectionCanvas.dispatchEvent(createPointerLikeMouseEvent("pointerdown", { clientX: hitX, clientY: hitY, button: 2 }));
    expect(app.getSelectedChartId()).toBe(chartId);

    selectionCanvas.dispatchEvent(createPointerLikeMouseEvent("pointerdown", { clientX: missX, clientY: missY, button: 2 }));
    expect(app.getSelectedChartId()).toBe(chartId);

    app.destroy();
    root.remove();
  });

  it("preserves ChartStore chart selection on context-click misses (with drawing interactions enabled)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("shared");
    expect((app as any).chartDrawingInteraction).toBeTruthy();

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Preserve selection (interactions enabled)",
      position: "A1",
    });
    // Deterministic absolute anchor for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(0),
      yEmu: pxToEmu(0),
      cxEmu: pxToEmu(100),
      cyEmu: pxToEmu(100),
    });

    const selectionCanvas = (app as any).selectionCanvas as HTMLCanvasElement;
    selectionCanvas.getBoundingClientRect = root.getBoundingClientRect as any;

    const rowHeaderWidth = (app as any).rowHeaderWidth as number;
    const colHeaderHeight = (app as any).colHeaderHeight as number;

    const hitX = rowHeaderWidth + 10;
    const hitY = colHeaderHeight + 10;
    const missX = rowHeaderWidth + 300;
    const missY = colHeaderHeight + 300;

    expect(app.hitTestDrawingAtClientPoint(missX, missY)).toBe(null);

    selectionCanvas.dispatchEvent(createPointerLikeMouseEvent("pointerdown", { clientX: hitX, clientY: hitY, button: 2 }));
    expect(app.getSelectedChartId()).toBe(chartId);

    const missDown = createPointerLikeMouseEvent("pointerdown", { clientX: missX, clientY: missY, button: 2 });
    selectionCanvas.dispatchEvent(missDown);
    expect(app.getSelectedChartId()).toBe(chartId);
    expect((missDown as any).__formulaDrawingContextClick).toBeUndefined();

    app.destroy();
    root.remove();
  });
});
