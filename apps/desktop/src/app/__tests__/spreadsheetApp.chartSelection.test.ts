/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;
let priorCanvasCharts: string | undefined;

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
  opts: { clientX: number; clientY: number; pointerId?: number; button?: number; pointerType?: string },
): void {
  const pointerId = opts.pointerId ?? 1;
  const button = opts.button ?? 0;
  const pointerType = opts.pointerType ?? "mouse";
  const base = { bubbles: true, cancelable: true, clientX: opts.clientX, clientY: opts.clientY, pointerId, button };
  const event =
    typeof (globalThis as any).PointerEvent === "function"
      ? new (globalThis as any).PointerEvent(type, { ...base, pointerType })
      : (() => {
          const e = new MouseEvent(type, base);
          Object.assign(e, { pointerId, pointerType });
          return e;
        })();
  // Some test environments polyfill `PointerEvent` as `MouseEvent` or omit `pointerId/pointerType`.
  // Ensure the fields exist so shared-grid and drawing interactions can identify context-clicks.
  Object.assign(event, { pointerId, pointerType });
  target.dispatchEvent(event);
}

describe("SpreadsheetApp chart selection + drag", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    // Default these tests to legacy chart mode; canvas charts mode is exercised via
    // the explicit `process.env.CANVAS_CHARTS = "1"` cases below.
    process.env.CANVAS_CHARTS = "0";
    document.body.innerHTML = "";

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

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("selects a chart on click", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
    const chart = app.listCharts().find((c) => c.id === chartId);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + rect.left + 2;
    const clickY = originY + rect.top + 2;
    dispatchPointerEvent(root, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 1 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 1 });

    expect(app.getSelectedChartId()).toBe(chart!.id);

    app.destroy();
    root.remove();
  });

  it("Escape clears chart selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
    const chart = app.listCharts().find((c) => c.id === chartId);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + rect.left + 2;
    const clickY = originY + rect.top + 2;
    dispatchPointerEvent(root, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 1 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 1 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(app.getSelectedChartId()).toBe(null);

    app.destroy();
    root.remove();
  });

  it("switching sheets clears chart selection even if the chart is on the target sheet", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(false);
      const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
      const chart = app.listCharts().find((c) => c.id === chartId);
      expect(chart).toBeTruthy();

      const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
      expect(rect).not.toBeNull();
      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const clickX = originX + rect.left + 2;
      const clickY = originY + rect.top + 2;
      dispatchPointerEvent(root, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 91 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 91 });
      expect(app.getSelectedChartId()).toBe(chart!.id);

      // Ensure the target sheet exists before switching.
      app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

      // Simulate moving the selected chart to the destination sheet (e.g. via cut/paste or remote update).
      const store = (app as any).chartStore as any;
      store.charts = store.charts.map((c: any) => (c.id === chart!.id ? { ...c, sheetId: "Sheet2" } : c));

      app.activateSheet("Sheet2");

      expect(app.getSelectedChartId()).toBe(null);
      expect(app.getSelectedDrawingId()).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("switching sheets mid-drag cancels an in-progress chart drag (legacy charts)", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(false);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Legacy Chart Sheet Switch Drag",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      const beforeAnchor = { ...(before!.anchor as any) };

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + 100;
      const endY = startY;

      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 401 });
      dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 401 });
      expect((app as any).chartDragState).not.toBeNull();

      const moved = app.listCharts().find((c) => c.id === result.chart_id);
      expect(moved).toBeTruthy();
      expect(moved!.anchor).not.toMatchObject(beforeAnchor);

      // Ensure the target sheet exists before switching.
      app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

      // Switch sheets while the pointer is still down. This should cancel the active chart drag
      // and revert the chart anchor to the initial pointerdown snapshot.
      app.activateSheet("Sheet2");
      expect((app as any).chartDragState).toBeNull();

      // Release the pointer after switching sheets (should be a no-op).
      dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 401 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor).toMatchObject(beforeAnchor);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("ignores pointerdown events from scrollbars (does not select/deselect charts)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Wide Chart",
      position: "A1:H10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // Pick a point near the right edge of the viewport so it overlaps with where scrollbars render.
    const clickX = 790;
    const clickY = originY + rect.top + 10;

    // Sanity: this point is inside the chart's hit region.
    const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
    expect(hit?.chart?.id).toBe(chart!.id);

    const thumb = (app as any).vScrollbarThumb as HTMLElement;
    expect(thumb).toBeTruthy();

    // Clicking on the scrollbar thumb should not select or deselect charts, even if a chart
    // extends underneath the scrollbar layer.
    dispatchPointerEvent(thumb, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 77 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 77 });
    expect(app.getSelectedChartId()).toBe(null);

    // Select via the grid surface, then ensure the scrollbar click does not deselect.
    dispatchPointerEvent(root, "pointerdown", { clientX: originX + rect.left + 10, clientY: originY + rect.top + 10, pointerId: 78 });
    dispatchPointerEvent(window, "pointerup", { clientX: originX + rect.left + 10, clientY: originY + rect.top + 10, pointerId: 78 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    dispatchPointerEvent(thumb, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 79 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 79 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    app.destroy();
    root.remove();
  });

  it("canvas charts mode: switching sheets clears chart selection even if the chart is on the target sheet", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart Sheet Switch Selection",
        position: "A1",
      });

      // Select the chart.
      (app as any).setSelectedChartId(result.chart_id);
      expect(app.getSelectedChartId()).toBe(result.chart_id);
      expect(app.getSelectedDrawingId()).toBeTruthy();

      // Ensure the target sheet exists before switching.
      app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

      // Simulate moving the selected chart to the destination sheet.
      const store = (app as any).chartStore as any;
      store.charts = store.charts.map((c: any) => (c.id === result.chart_id ? { ...c, sheetId: "Sheet2" } : c));

      app.activateSheet("Sheet2");

      expect(app.getSelectedChartId()).toBe(null);
      expect(app.getSelectedDrawingId()).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: ignores pointerdown events from scrollbars (does not select/deselect charts)", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Wide Canvas Chart",
        position: "A1:H10",
      });
      const chart = app.listCharts().find((c) => c.id === result.chart_id);
      expect(chart).toBeTruthy();

      const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
      expect(rect).not.toBeNull();

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      // Pick a point near the right edge of the viewport so it overlaps with where scrollbars render.
      const clickX = 790;
      const clickY = originY + rect.top + 10;

      // Sanity: this point is inside the chart's hit region.
      const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
      expect(hit?.chart?.id).toBe(chart!.id);

      const thumb = (app as any).vScrollbarThumb as HTMLElement;
      expect(thumb).toBeTruthy();

      // Clicking on the scrollbar thumb should not select charts, even if a chart extends underneath.
      dispatchPointerEvent(thumb, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 81 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 81 });
      expect(app.getSelectedChartId()).toBe(null);

      // Select via the grid surface, then ensure the scrollbar click does not deselect.
      dispatchPointerEvent(root, "pointerdown", { clientX: originX + rect.left + 10, clientY: originY + rect.top + 10, pointerId: 82 });
      dispatchPointerEvent(window, "pointerup", { clientX: originX + rect.left + 10, clientY: originY + rect.top + 10, pointerId: 82 });
      expect(app.getSelectedChartId()).toBe(chart!.id);

      dispatchPointerEvent(thumb, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 83 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 83 });
      expect(app.getSelectedChartId()).toBe(chart!.id);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("context-click selects a chart without moving the active cell", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Move the active cell somewhere else so we can verify the context-click does not change it.
    app.selectRange({ range: { startRow: 2, endRow: 2, startCol: 2, endCol: 2 } }, { scrollIntoView: false, focus: false });
    expect(status.activeCell.textContent).toBe("C3");

    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();
    // Shared-grid selection uses the selection canvas's client rect to map pointer coords.
    selectionCanvas!.getBoundingClientRect = root.getBoundingClientRect as any;

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Context Click Chart",
      position: "A1:H10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + rect.left + 10;
    const clickY = originY + rect.top + 10;

    // Sanity: this point is inside the chart's hit region.
    const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
    expect(hit?.chart?.id).toBe(chart!.id);

    // Context click should select the chart, but should not move the active cell underneath.
    dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 90, button: 2 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 90, button: 2 });

    expect(app.getSelectedChartId()).toBe(chart!.id);
    expect(status.activeCell.textContent).toBe("C3");

    app.destroy();
    root.remove();
  });

  it("context-click on a drawing does not select a chart underneath when drawing interactions are enabled", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart Under Drawing",
      position: "A1:H10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    // Inject a simple drawing that overlaps the chart area (absolute coords are relative to A1, under headers).
    const doc = (app as any).document as any;
    doc.getSheetDrawings = () => [
      {
        id: 1,
        kind: { type: "shape", label: "Overlay" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(40), yEmu: pxToEmu(40) },
          size: { cx: pxToEmu(120), cy: pxToEmu(80) },
        },
        zOrder: 0,
      },
    ];
    (app as any).drawingObjectsCache = null;

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + 60;
    const clickY = originY + 60;

    // Sanity: the chart hit test sees the chart at this point.
    const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
    expect(hit?.chart?.id).toBe(chart!.id);

    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();
    selectionCanvas!.getBoundingClientRect = root.getBoundingClientRect as any;

    dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 91, button: 2 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 91, button: 2 });

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getSelectedChartId()).toBe(null);

    app.destroy();
    root.remove();
  });

  it("when a drawing overlaps a chart, clicking selects the drawing (not the chart)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart Under Drawing",
      position: "A1:H10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    // Inject a simple drawing that overlaps the chart area (absolute coords are relative to A1, under headers).
    const doc = (app as any).document as any;
    doc.getSheetDrawings = () => [
      {
        id: 1,
        kind: { type: "shape", label: "Overlay" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(40), yEmu: pxToEmu(40) },
          size: { cx: pxToEmu(120), cy: pxToEmu(80) },
        },
        zOrder: 0,
      },
    ];
    (app as any).drawingObjectsCache = null;

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const clickX = originX + 60;
    const clickY = originY + 60;

    // Sanity: the chart hit test sees the chart at this point.
    const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
    expect(hit?.chart?.id).toBe(chart!.id);

    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();

    dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 92 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 92 });

    expect(app.getSelectedDrawingId()).toBe(1);
    expect(app.getSelectedChartId()).toBe(null);

    app.destroy();
    root.remove();
  });

  it("chart resize handles win over overlapping drawings when the chart is selected", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Resizable Chart",
      // Keep the chart fully within the visible cell area so the bottom-right resize handle
      // is hit-testable in this 800x600 test viewport (shared-grid headers reduce the cell area).
      position: "A1:G10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // Select the chart first so resize handles are active.
    const selectX = originX + rect.left + 10;
    const selectY = originY + rect.top + 10;
    dispatchPointerEvent(root, "pointerdown", { clientX: selectX, clientY: selectY, pointerId: 120 });
    dispatchPointerEvent(window, "pointerup", { clientX: selectX, clientY: selectY, pointerId: 120 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    // Add a drawing that overlaps the bottom-right resize handle.
    const doc = (app as any).document as any;
    doc.getSheetDrawings = () => [
      {
        id: 1,
        kind: { type: "shape", label: "Overlay" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(2000), cy: pxToEmu(2000) },
        },
        zOrder: 0,
      },
    ];
    (app as any).drawingObjectsCache = null;

    const handleX = originX + rect.left + rect.width;
    const handleY = originY + rect.top + rect.height;

    dispatchPointerEvent(root, "pointerdown", { clientX: handleX, clientY: handleY, pointerId: 121 });
    expect((app as any).chartDragState?.mode).toBe("resize");
    dispatchPointerEvent(window, "pointerup", { clientX: handleX, clientY: handleY, pointerId: 121 });

    expect(app.getSelectedChartId()).toBe(chart!.id);
    expect((app as any).selectedDrawingId).toBe(null);

    app.destroy();
    root.remove();
  });

  it("does not let clipped drawing handles (frozen-pane mismatch) block chart selection", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Disable the dedicated drawing interaction controller so `onChartPointerDownCapture` must
    // correctly arbitrate between charts and (selected) drawing selection handles.
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });

    // Freeze the first column so drawings anchored in the scrollable pane are clipped away from the
    // frozen pane (including selection handles).
    const docAny = (app as any).document as any;
    docAny.setFrozen?.(app.getCurrentSheetId(), 0, 1, { label: "Freeze first column" });
    (app as any).syncFrozenPanes?.();
    expect(app.getFrozen().frozenCols).toBe(1);

    const viewport = app.getDrawingInteractionViewport();
    const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
    const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
    const frozenBoundaryX = Number.isFinite(viewport.frozenWidthPx) ? (viewport.frozenWidthPx as number) : headerOffsetX;

    // Place a drawing so its top-left handle lies in the frozen pane, but the drawing itself
    // belongs to the scrollable pane (absolute anchors always scroll).
    const posX = viewport.scrollX + (frozenBoundaryX - headerOffsetX) - 10;
    const posY = viewport.scrollY + 80;
    app.setDrawingObjects([
      {
        id: 1,
        kind: { type: "shape", label: "Clipped" },
        anchor: {
          type: "absolute",
          pos: { xEmu: pxToEmu(posX), yEmu: pxToEmu(posY) },
          size: { cx: pxToEmu(50), cy: pxToEmu(40) },
        },
        zOrder: 0,
      },
    ]);
    app.selectDrawingById(1);

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart Under Clipped Handle",
      position: "A1:H10",
    });
    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const clickX = frozenBoundaryX - 10;
    const clickY = headerOffsetY + (posY - viewport.scrollY);

    // Sanity: the drawing hit-test should treat this as a miss (it's clipped out of the active pane),
    // but the chart hit-test should see the chart at this coordinate.
    expect(app.hitTestDrawingAtClientPoint(clickX, clickY)).toBeNull();
    expect((app as any).hitTestChartAtClientPoint(clickX, clickY)?.chart?.id).toBe(chart!.id);

    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();

    dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 122 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 122 });

    expect(app.getSelectedChartId()).toBe(chart!.id);
    expect(app.getSelectedDrawingId()).toBe(null);

    app.destroy();
    root.remove();
  });

  it("canvas charts mode: when a drawing overlaps a chart, clicking selects the chart (charts are above drawings)", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Chart Under Drawing",
        position: "A1:H10",
      });
      const chart = app.listCharts().find((c) => c.id === result.chart_id);
      expect(chart).toBeTruthy();

      // Inject a simple drawing that overlaps the chart area.
      const doc = (app as any).document as any;
      doc.getSheetDrawings = () => [
        {
          id: 1,
          kind: { type: "shape", label: "Overlay" },
          anchor: {
            type: "absolute",
            pos: { xEmu: pxToEmu(40), yEmu: pxToEmu(40) },
            size: { cx: pxToEmu(120), cy: pxToEmu(80) },
          },
          zOrder: 0,
        },
      ];
      (app as any).drawingObjectsCache = null;

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const clickX = originX + 60;
      const clickY = originY + 60;

      // Sanity: chart hit test sees a chart at this point.
      const hit = (app as any).hitTestChartAtClientPoint(clickX, clickY);
      expect(hit?.chart?.id).toBe(chart!.id);

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();
      selectionCanvas!.getBoundingClientRect = root.getBoundingClientRect as any;

      dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 93 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 93 });

      expect(app.getSelectedChartId()).toBe(chart!.id);
      expect((app as any).selectedDrawingId).toBe(null);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("while formula bar is editing a formula, clicking a chart does not select it", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const formulaBar = document.createElement("div");
    document.body.appendChild(formulaBar);

    const app = new SpreadsheetApp(root, status, { formulaBar });
    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
    const chart = app.listCharts().find((c) => c.id === chartId);
    expect(chart).toBeTruthy();

    const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // Force formula-bar state to "formula editing" so SpreadsheetApp should ignore chart pointerdowns.
    const fb = (app as any).formulaBar as { model: { updateDraft: (draft: string, start: number, end: number) => void }; isFormulaEditing: () => boolean };
    fb.model.updateDraft("=A1", 3, 3);
    expect(fb.isFormulaEditing()).toBe(true);

    const clickX = originX + rect.left + 2;
    const clickY = originY + rect.top + 2;

    const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    expect(selectionCanvas).not.toBeNull();

    dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 100 });
    dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 100 });

    expect(app.getSelectedChartId()).toBe(null);

    app.destroy();
    root.remove();
    formulaBar.remove();
  });

  it("canvas charts mode: while formula bar is editing a formula, clicking a chart does not select it", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const formulaBar = document.createElement("div");
      document.body.appendChild(formulaBar);

      const app = new SpreadsheetApp(root, status, { formulaBar });
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Formula Editing Canvas Chart",
        position: "A1:H10",
      });
      const chart = app.listCharts().find((c) => c.id === result.chart_id);
      expect(chart).toBeTruthy();

      const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
      expect(rect).not.toBeNull();

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const fb = (app as any).formulaBar as { model: { updateDraft: (draft: string, start: number, end: number) => void }; isFormulaEditing: () => boolean };
      fb.model.updateDraft("=A1", 3, 3);
      expect(fb.isFormulaEditing()).toBe(true);

      const clickX = originX + rect.left + 2;
      const clickY = originY + rect.top + 2;

      const selectionCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
      expect(selectionCanvas).not.toBeNull();

      dispatchPointerEvent(selectionCanvas!, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 101 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 101 });

      expect(app.getSelectedChartId()).toBe(null);

      app.destroy();
      root.remove();
      formulaBar.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("dragging a chart updates its twoCell anchor", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Drag Chart",
      position: "A1",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");

    const beforeAnchor = before!.anchor as any;
    expect(beforeAnchor.fromCol).toBe(0);
    expect(beforeAnchor.toCol).toBeGreaterThan(0);

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const startX = originX + rect.left + 10;
    const startY = originY + rect.top + 10;
    const endX = startX + 100; // move by one column (default col width)
    const endY = startY;

    dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 7 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 7 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 7 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");

    const afterAnchor = after!.anchor as any;
    expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol + 1);
    expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
    expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
    expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

    app.destroy();
    root.remove();
  });

  it("dragging a frozen chart into the scrollable pane accounts for scroll offsets", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setFrozen(app.getCurrentSheetId(), 0, 1, { label: "Freeze Col" });
    app.setScroll(200, 0);

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Frozen Drag",
      position: "A1",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // Drag chart so its top-left ends up at x=150px in the cell-area (past the frozen boundary).
    const startX = originX + rect.left + 10;
    const startY = originY + rect.top + 10;
    const endX = startX + (150 - rect.left);
    const endY = startY;

    dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 55 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 55 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 55 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");
    const afterAnchor = after!.anchor as any;

    // With scrollX=200, landing at x=150 should map to sheet-space x=350 => col 3.
    expect(afterAnchor.fromCol).toBe(3);
    expect(afterAnchor.toCol).toBe(8);

    app.destroy();
    root.remove();
  });

  it("dragging a scrollable chart whose top-left is offscreen does not snap to frozen pane", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setFrozen(app.getCurrentSheetId(), 0, 1, { label: "Freeze Col" });
    app.setScroll(400, 0);

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Offscreen Scrollable",
      position: "D1:H10",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");
    const beforeAnchor = before!.anchor as any;
    expect(beforeAnchor.fromCol).toBe(3);
    expect(beforeAnchor.toCol).toBe(8);

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();
    // The chart's top-left is offscreen to the left (negative) due to scroll, but it is still
    // partially visible in the scrollable pane due to clipping.
    expect(rect.left).toBeLessThan(0);

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    const startX = originX + 150;
    const startY = originY + rect.top + 10;
    const endX = startX + 100;
    const endY = startY;

    dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 56 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 56 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 56 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");
    const afterAnchor = after!.anchor as any;

    // Dragging right by one column should shift the anchor by one column, not snap to frozen col 0.
    expect(afterAnchor.fromCol).toBe(4);
    expect(afterAnchor.toCol).toBe(9);

    app.destroy();
    root.remove();
  });

  it("resizing a selected chart updates its twoCell anchor", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Resize Chart",
      position: "A1",
    });

    const before = app.listCharts().find((c) => c.id === result.chart_id);
    expect(before).toBeTruthy();
    expect(before!.anchor.kind).toBe("twoCell");

    const beforeAnchor = before!.anchor as any;
    expect(beforeAnchor.fromCol).toBe(0);
    expect(beforeAnchor.fromRow).toBe(0);

    const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
    expect(rect).not.toBeNull();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // First click selects the chart.
    const selectX = originX + rect.left + 10;
    const selectY = originY + rect.top + 10;
    dispatchPointerEvent(root, "pointerdown", { clientX: selectX, clientY: selectY, pointerId: 31 });
    dispatchPointerEvent(window, "pointerup", { clientX: selectX, clientY: selectY, pointerId: 31 });
    expect(app.getSelectedChartId()).toBe(result.chart_id);

    // Second click on the bottom-right handle starts a resize drag.
    const handleX = originX + rect.left + rect.width;
    const handleY = originY + rect.top + rect.height;
    const endX = handleX + 110; // increase width by ~1 column (default col width = 100)
    const endY = handleY;

    dispatchPointerEvent(root, "pointerdown", { clientX: handleX, clientY: handleY, pointerId: 32 });
    dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 32 });
    dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 32 });

    const after = app.listCharts().find((c) => c.id === result.chart_id);
    expect(after).toBeTruthy();
    expect(after!.anchor.kind).toBe("twoCell");

    const afterAnchor = after!.anchor as any;
    expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol);
    expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
    expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
    expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

    app.destroy();
    root.remove();
  });

  it("hit testing respects shared-grid pane clipping (frozen panes)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

    const result = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Frozen Chart",
      position: "A1",
    });

    const chart = app.listCharts().find((c) => c.id === result.chart_id);
    expect(chart).toBeTruthy();

    const layout = (app as any).chartOverlayLayout();
    const originX = layout.originX as number;
    const originY = layout.originY as number;

    // This point is inside the chart bounds (which extend far beyond A1), but outside the
    // top-left frozen pane (so it should not count as a hit because the chart is clipped).
    dispatchPointerEvent(root, "pointerdown", { clientX: originX + 150, clientY: originY + 10, pointerId: 99 });
    dispatchPointerEvent(window, "pointerup", { clientX: originX + 150, clientY: originY + 10, pointerId: 99 });
    expect(app.getSelectedChartId()).toBe(null);

    // This point lies in the visible (clipped) portion of the chart in the top-left pane.
    dispatchPointerEvent(root, "pointerdown", { clientX: originX + 50, clientY: originY + 10, pointerId: 100 });
    dispatchPointerEvent(window, "pointerup", { clientX: originX + 50, clientY: originY + 10, pointerId: 100 });
    expect(app.getSelectedChartId()).toBe(chart!.id);

    app.destroy();
    root.remove();
  });

  it("canvas charts mode: dragging a chart updates its anchor", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart Drag",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      expect(before!.anchor.kind).toBe("twoCell");
      const beforeAnchor = before!.anchor as any;

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + 100; // move by one column (default col width)
      const endY = startY;

      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 201 });
      dispatchPointerEvent(root, "pointermove", { clientX: endX, clientY: endY, pointerId: 201 });
      dispatchPointerEvent(root, "pointerup", { clientX: endX, clientY: endY, pointerId: 201 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor.kind).toBe("twoCell");
      const afterAnchor = after!.anchor as any;

      expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol + 1);
      expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
      expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
      expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: switching sheets mid-drag cancels the chart gesture (no stale gesture state)", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };
 
      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart Sheet Switch Drag",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      const beforeAnchor = { ...(before!.anchor as any) };

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + 100; // move by one column
      const endY = startY;

      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 301 });
      dispatchPointerEvent(root, "pointermove", { clientX: endX, clientY: endY, pointerId: 301 });
      expect((app as any).chartDrawingGestureActive).toBe(true);

      // Ensure the target sheet exists before switching.
      app.getDocument().setCellValue("Sheet2", { row: 0, col: 0 }, "X");

      // Switch sheets while the pointer is still down. This should cancel the active chart drag.
      app.activateSheet("Sheet2");
      expect((app as any).chartDrawingGestureActive).toBe(false);

      // Release the pointer after switching sheets (should be a no-op).
      dispatchPointerEvent(root, "pointerup", { clientX: endX, clientY: endY, pointerId: 301 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor).toMatchObject(beforeAnchor);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: Escape cancels an in-progress chart drag without clearing selection", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart Escape Drag",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      const beforeAnchor = { ...(before!.anchor as any) };

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + 100; // move by one column
      const endY = startY;

      // Start a drag gesture (do not pointerup yet).
      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 301 });
      expect(app.getSelectedChartId()).toBe(result.chart_id);

      dispatchPointerEvent(root, "pointermove", { clientX: endX, clientY: endY, pointerId: 301 });
      // Canvas charts use DrawingInteractionController to provide live preview during a drag,
      // but do not commit the new anchor into `ChartStore` until pointerup. The live state is
      // exposed via SpreadsheetApp's in-flight override list.
      expect((app as any).chartDrawingGestureActive).toBe(true);
      const override = (app as any).canvasChartDrawingObjectsOverride;
      expect(override?.sheetId).toBe(app.getCurrentSheetId());
      const previewObj = override?.objects?.find?.((obj: any) => obj?.kind?.type === "chart" && obj?.kind?.chartId === result.chart_id);
      expect(previewObj).toBeTruthy();
      expect((previewObj!.anchor as any)?.from?.cell?.col).toBe((beforeAnchor as any).fromCol + 1);
      const moved = app.listCharts().find((c) => c.id === result.chart_id);
      expect(moved).toBeTruthy();
      expect(moved!.anchor).toEqual(beforeAnchor);

      // Press Escape during the drag: should cancel the gesture and keep the chart selected.
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      expect((app as any).chartDrawingGestureActive).toBe(false);

      const afterEscape = app.listCharts().find((c) => c.id === result.chart_id);
      expect(afterEscape).toBeTruthy();
      expect(afterEscape!.anchor).toEqual(beforeAnchor);
      expect(app.getSelectedChartId()).toBe(result.chart_id);

      // Clean up any pending pointer state.
      dispatchPointerEvent(root, "pointerup", { clientX: endX, clientY: endY, pointerId: 301 });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: resizing a chart updates its twoCell anchor", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart Resize",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      expect(before!.anchor.kind).toBe("twoCell");
      const beforeAnchor = before!.anchor as any;

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      // Start the resize drag by grabbing the SE handle.
      const handleX = originX + rect.left + rect.width;
      const handleY = originY + rect.top + rect.height;
      const endX = handleX + 110; // increase width by ~1 column (default col width = 100)
      const endY = handleY;

      dispatchPointerEvent(root, "pointerdown", { clientX: handleX, clientY: handleY, pointerId: 202 });
      dispatchPointerEvent(root, "pointermove", { clientX: endX, clientY: endY, pointerId: 202 });
      dispatchPointerEvent(root, "pointerup", { clientX: endX, clientY: endY, pointerId: 202 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor.kind).toBe("twoCell");
      const afterAnchor = after!.anchor as any;

      expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol);
      expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
      expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
      expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: hit testing respects shared-grid pane clipping (frozen panes)", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Frozen Hit Test",
        position: "A1",
      });

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      // This point is inside the chart bounds (it extends far beyond A1), but outside the
      // top-left frozen pane (so it should not count as a hit because the chart is clipped).
      dispatchPointerEvent(root, "pointerdown", { clientX: originX + 150, clientY: originY + 10, pointerId: 203 });
      dispatchPointerEvent(root, "pointerup", { clientX: originX + 150, clientY: originY + 10, pointerId: 203 });
      expect(app.getSelectedChartId()).toBe(null);

      // This point lies in the visible (clipped) portion of the chart in the top-left pane.
      dispatchPointerEvent(root, "pointerdown", { clientX: originX + 50, clientY: originY + 10, pointerId: 204 });
      dispatchPointerEvent(root, "pointerup", { clientX: originX + 50, clientY: originY + 10, pointerId: 204 });
      expect(app.getSelectedChartId()).toBe(result.chart_id);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });

  it("canvas charts mode: dragging a frozen chart into the scrollable pane accounts for scroll offsets", () => {
    const prior = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "1";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect((app as any).useCanvasCharts).toBe(true);

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 0, 1, { label: "Freeze Col" });
      app.setScroll(200, 0);

      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Frozen Drag",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      expect(before!.anchor.kind).toBe("twoCell");

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const viewport = app.getDrawingInteractionViewport();
      const originX = (viewport.headerOffsetX ?? 0) as number;
      const originY = (viewport.headerOffsetY ?? 0) as number;

      // Drag chart so its top-left ends up at x=150px in the cell-area (past the frozen boundary).
      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + (150 - rect.left);
      const endY = startY;

      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 205 });
      dispatchPointerEvent(root, "pointermove", { clientX: endX, clientY: endY, pointerId: 205 });
      dispatchPointerEvent(root, "pointerup", { clientX: endX, clientY: endY, pointerId: 205 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor.kind).toBe("twoCell");
      const afterAnchor = after!.anchor as any;

      // With scrollX=200, landing at x=150 should map to sheet-space x=350 => col 3.
      expect(afterAnchor.fromCol).toBe(3);
      expect(afterAnchor.toCol).toBe(8);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = prior;
    }
  });
});

describe("SpreadsheetApp chart selection + drag (legacy grid)", () => {
  it("selects a chart on click", () => {
    const priorCharts = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.DESKTOP_GRID_MODE = "legacy";
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    try {
      const app = new SpreadsheetApp(root, status);
      const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
      const chart = app.listCharts().find((c) => c.id === chartId);
      expect(chart).toBeTruthy();

      const rect = (app as any).chartAnchorToViewportRect(chart!.anchor);
      expect(rect).not.toBeNull();

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const clickX = originX + rect.left + 2;
      const clickY = originY + rect.top + 2;
      dispatchPointerEvent(root, "pointerdown", { clientX: clickX, clientY: clickY, pointerId: 1 });
      dispatchPointerEvent(window, "pointerup", { clientX: clickX, clientY: clickY, pointerId: 1 });

      expect(app.getSelectedChartId()).toBe(chart!.id);

      app.destroy();
      root.remove();
    } finally {
      if (priorCharts === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = priorCharts;
    }
  });

  it("dragging a chart updates its twoCell anchor", () => {
    const priorCharts = process.env.CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.DESKTOP_GRID_MODE = "legacy";
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    try {
      const app = new SpreadsheetApp(root, status);
      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Drag Chart Legacy",
        position: "A1",
      });

      const before = app.listCharts().find((c) => c.id === result.chart_id);
      expect(before).toBeTruthy();
      expect(before!.anchor.kind).toBe("twoCell");

      const beforeAnchor = before!.anchor as any;
      expect(beforeAnchor.fromCol).toBe(0);
      expect(beforeAnchor.toCol).toBeGreaterThan(0);

      const rect = (app as any).chartAnchorToViewportRect(before!.anchor);
      expect(rect).not.toBeNull();

      const layout = (app as any).chartOverlayLayout();
      const originX = layout.originX as number;
      const originY = layout.originY as number;

      const startX = originX + rect.left + 10;
      const startY = originY + rect.top + 10;
      const endX = startX + 100; // move by one column (default col width)
      const endY = startY;

      dispatchPointerEvent(root, "pointerdown", { clientX: startX, clientY: startY, pointerId: 7 });
      dispatchPointerEvent(window, "pointermove", { clientX: endX, clientY: endY, pointerId: 7 });
      dispatchPointerEvent(window, "pointerup", { clientX: endX, clientY: endY, pointerId: 7 });

      const after = app.listCharts().find((c) => c.id === result.chart_id);
      expect(after).toBeTruthy();
      expect(after!.anchor.kind).toBe("twoCell");

      const afterAnchor = after!.anchor as any;
      expect(afterAnchor.fromCol).toBe(beforeAnchor.fromCol + 1);
      expect(afterAnchor.toCol).toBe(beforeAnchor.toCol + 1);
      expect(afterAnchor.fromRow).toBe(beforeAnchor.fromRow);
      expect(afterAnchor.toRow).toBe(beforeAnchor.toRow);

      app.destroy();
      root.remove();
    } finally {
      if (priorCharts === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = priorCharts;
    }
  });
});
