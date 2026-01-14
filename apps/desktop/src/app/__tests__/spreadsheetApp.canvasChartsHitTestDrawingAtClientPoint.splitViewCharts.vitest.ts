/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../../grid/splitView/secondaryGridView.js";
import { drawingObjectToViewportRect } from "../../drawings/hitTest";
import { pxToEmu } from "../../drawings/overlay";
import type { DrawingObject } from "../../drawings/types";
import { chartIdToDrawingId } from "../../charts/chartDrawingAdapter";
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

describe("SpreadsheetApp.hitTestDrawingAtClientPoint (canvas charts, split view)", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    process.env.CANVAS_CHARTS = "1";

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

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("hit tests ChartStore charts inside the secondary split-view pane", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(app.getCurrentSheetId()).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    expect(app.hitTestDrawingAtClientPoint(hitClientX, hitClientY)).toEqual({ id: chartDrawingId });

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("moves ChartStore charts via drag gestures in the secondary split-view pane", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const beforeObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(beforeObj).toBeTruthy();
    expect(beforeObj!.anchor.type).toBe("absolute");
    const beforeAnchor = beforeObj!.anchor as any;

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(beforeObj!, viewport as any, geom as any);
    const startX = objRect.x + 5;
    const startY = objRect.y + 5;
    const dx = 30;
    const dy = 20;
    const endX = startX + dx;
    const endY = startY + dy;

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", { clientX: rect.left + startX, clientY: rect.top + startY, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", { clientX: rect.left + endX, clientY: rect.top + endY, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: rect.left + endX, clientY: rect.top + endY, button: 0, pointerId: 1 }),
    );

    const afterObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(afterObj).toBeTruthy();
    expect(afterObj!.anchor.type).toBe("absolute");
    const afterAnchor = afterObj!.anchor as any;

    expect(afterAnchor.pos.xEmu).toBeCloseTo(beforeAnchor.pos.xEmu + pxToEmu(dx, zoom), 4);
    expect(afterAnchor.pos.yEmu).toBeCloseTo(beforeAnchor.pos.yEmu + pxToEmu(dy, zoom), 4);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("right-click selects ChartStore charts in the secondary pane without moving the grid selection", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    const beforeSelection = view.grid.renderer.getSelection();
    const beforeSelectionCopy = beforeSelection ? { ...beforeSelection } : beforeSelection;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: hitClientX,
      clientY: hitClientY,
      button: 2,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(down);

    // DrawingInteractionController should tag the event so DesktopSharedGrid does not move the active cell
    // underneath the chart (Excel-like behavior).
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(app.getSelectedChartId()).toBe(chartId);

    const afterSelection = view.grid.renderer.getSelection();
    expect(afterSelection).toEqual(beforeSelectionCopy);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("preserves ChartStore chart selection on context-click misses in the secondary pane", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    const missClientX = rect.left + rect.width - 10;
    const missClientY = rect.top + rect.height - 10;
    expect(app.hitTestDrawingAtClientPoint(missClientX, missClientY)).toBe(null);

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: hitClientX,
        clientY: hitClientY,
        button: 2,
        pointerId: 1,
      }),
    );
    expect(app.getSelectedChartId()).toBe(chartId);

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: missClientX,
        clientY: missClientY,
        button: 2,
        pointerId: 2,
      }),
    );
    expect(app.getSelectedChartId()).toBe(chartId);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("deletes ChartStore charts via Delete key from the secondary pane", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    // Select the chart from the secondary pane.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", { clientX: hitClientX, clientY: hitClientY, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: hitClientX, clientY: hitClientY, button: 0, pointerId: 1 }),
    );
    expect(app.getSelectedChartId()).toBe(chartId);
    const chartCountBefore = app.listCharts().length;
    expect(chartCountBefore).toBeGreaterThan(0);

    // Delete should remove the selected chart, even though focus is in the secondary pane.
    const del = new KeyboardEvent("keydown", { key: "Delete", bubbles: true, cancelable: true });
    selectionCanvas.dispatchEvent(del);
    expect(del.defaultPrevented).toBe(true);
    expect(app.getSelectedChartId()).toBeNull();
    expect(app.listCharts().some((c) => c.id === chartId)).toBe(false);
    expect(app.listCharts()).toHaveLength(chartCountBefore - 1);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("clears ChartStore chart selection via Escape key from the secondary pane", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    // Select the chart from the secondary pane.
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", { clientX: hitClientX, clientY: hitClientY, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: hitClientX, clientY: hitClientY, button: 0, pointerId: 1 }),
    );
    expect(app.getSelectedChartId()).toBe(chartId);
    const chartCountBefore = app.listCharts().length;

    const esc = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    selectionCanvas.dispatchEvent(esc);
    expect(esc.defaultPrevented).toBe(true);
    expect(app.getSelectedChartId()).toBeNull();
    expect(app.listCharts()).toHaveLength(chartCountBefore);
    expect(app.listCharts().some((c) => c.id === chartId)).toBe(true);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("right-click selects a ChartStore chart over an overlapping drawing when drawing interactions are enabled", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    // Enable drawing interactions so we exercise the split-view secondary pane's dedicated
    // DrawingInteractionController (which needs to yield to charts when they overlap).
    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Override anchor to a deterministic absolute position for stable hit testing.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(30),
      yEmu: pxToEmu(20),
      cxEmu: pxToEmu(80),
      cyEmu: pxToEmu(60),
    });

    // Add a drawing underneath the chart so we can verify z-order arbitration (chart should win).
    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-under-chart" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(20), yEmu: pxToEmu(10) },
        size: { cx: pxToEmu(120), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };
    app.setDrawingObjects([drawing]);

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => app.getCurrentSheetId(),
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    const beforeSelection = view.grid.renderer.getSelection();
    const beforeSelectionCopy = beforeSelection ? { ...beforeSelection } : beforeSelection;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const sheetId = app.getCurrentSheetId();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const objRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);
    const hitClientX = rect.left + objRect.x + 5;
    const hitClientY = rect.top + objRect.y + 5;

    const down = createPointerLikeMouseEvent("pointerdown", {
      clientX: hitClientX,
      clientY: hitClientY,
      button: 2,
      pointerId: 1,
    });
    selectionCanvas.dispatchEvent(down);

    // Chart z-order should win, even though a drawing overlaps underneath.
    expect(app.getSelectedChartId()).toBe(chartId);
    expect((app as any).selectedDrawingId).toBeNull();
    expect(app.getSelectedDrawingId()).toBe(chartDrawingId);

    // DrawingInteractionController should tag the event so DesktopSharedGrid does not move the active cell
    // underneath the chart (Excel-like behavior).
    expect((down as any).__formulaDrawingContextClick).toBe(true);
    expect(down.defaultPrevented).toBe(false);
    expect(view.grid.renderer.getSelection()).toEqual(beforeSelectionCopy);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("allows resizing a selected drawing via its handles even when a chart overlaps (secondary pane)", () => {
    const root = document.createElement("div");
    root.tabIndex = 0;
    root.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 0,
          top: 0,
          right: 400,
          bottom: 300,
          x: 0,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(root);

    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: true });
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Split View Chart",
      position: "A1",
    });

    // Make the chart large so it overlaps the drawing's resize handles.
    (app as any).chartStore.updateChartAnchor(chartId, {
      kind: "absolute",
      xEmu: pxToEmu(20),
      yEmu: pxToEmu(10),
      cxEmu: pxToEmu(200),
      cyEmu: pxToEmu(150),
    });
    const chartBefore = app.listCharts().find((c) => c.id === chartId) ?? null;
    expect(chartBefore).toBeTruthy();
    const chartAnchorBefore = JSON.parse(JSON.stringify(chartBefore!.anchor));

    const drawing: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img-resize-under-chart" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(60), yEmu: pxToEmu(50) },
        size: { cx: pxToEmu(40), cy: pxToEmu(30) },
      },
      zOrder: 0,
    };
    const sheetId = app.getCurrentSheetId();
    (app.getDocument() as any).setSheetDrawings(sheetId, [drawing], { label: "Set Drawings" });
    app.syncSheetDrawings();

    const secondaryContainer = document.createElement("div");
    secondaryContainer.tabIndex = 0;
    Object.defineProperty(secondaryContainer, "clientWidth", { configurable: true, value: 400 });
    Object.defineProperty(secondaryContainer, "clientHeight", { configurable: true, value: 300 });
    secondaryContainer.getBoundingClientRect = vi.fn(
      () =>
        ({
          width: 400,
          height: 300,
          left: 500,
          top: 0,
          right: 900,
          bottom: 300,
          x: 500,
          y: 0,
          toJSON: () => {},
        }) as any,
    );
    document.body.appendChild(secondaryContainer);

    const view = new SecondaryGridView({
      container: secondaryContainer,
      document: app.getDocument(),
      getSheetId: () => sheetId,
      rowCount: 50,
      colCount: 50,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: (sheetId) => app.getDrawingObjects(sheetId),
      images: app.getDrawingImages(),
      getSelectedDrawingId: () => app.getSelectedDrawingId(),
      chartRenderer: app.getDrawingChartRenderer(),
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");
    selectionCanvas.getBoundingClientRect = secondaryContainer.getBoundingClientRect as any;

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    // Select the drawing first (e.g. via selection pane). Resize should still work even if the chart overlaps the handles.
    app.selectDrawingById(drawing.id);
    expect(app.getSelectedDrawingId()).toBe(drawing.id);
    expect(app.getSelectedChartId()).toBeNull();

    const drawingObj = app.getDrawingObjects(sheetId).find((o) => o.id === drawing.id) ?? null;
    expect(drawingObj).toBeTruthy();

    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = app.getDrawingObjects(sheetId).find((o) => o.id === chartDrawingId) ?? null;
    expect(chartObj).toBeTruthy();

    const rect = secondaryContainer.getBoundingClientRect();
    const scroll = view.grid.getScroll();
    const viewportState = view.grid.renderer.scroll.getViewportState();
    const headerRows = 1;
    const headerCols = 1;
    const headerWidth = view.grid.renderer.scroll.cols.totalSize(headerCols);
    const headerHeight = view.grid.renderer.scroll.rows.totalSize(headerRows);
    const headerOffsetX = Math.min(headerWidth, rect.width);
    const headerOffsetY = Math.min(headerHeight, rect.height);
    const zoom = view.grid.renderer.getZoom();
    const { frozenRows, frozenCols } = app.getFrozen();

    const viewport = {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width: rect.width,
      height: rect.height,
      dpr: 1,
      zoom,
      frozenRows,
      frozenCols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewportState.frozenWidth,
      frozenHeightPx: viewportState.frozenHeight,
    };

    const geom = {
      cellOriginPx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: view.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: view.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell: { row: number; col: number }) => {
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          width: view.grid.renderer.getColWidth(gridCol),
          height: view.grid.renderer.getRowHeight(gridRow),
        };
      },
    };

    const drawRect = drawingObjectToViewportRect(drawingObj!, viewport as any, geom as any);
    const chartRect = drawingObjectToViewportRect(chartObj!, viewport as any, geom as any);

    // Hit the drawing's bottom-right resize handle, which lies inside the chart bounds.
    const startX = drawRect.x + drawRect.width;
    const startY = drawRect.y + drawRect.height;
    // Be tolerant of minor float drift from EMU<->px conversions.
    expect(startX).toBeGreaterThanOrEqual(chartRect.x - 0.5);
    expect(startY).toBeGreaterThanOrEqual(chartRect.y - 0.5);
    expect(startX).toBeLessThanOrEqual(chartRect.x + chartRect.width + 0.5);
    expect(startY).toBeLessThanOrEqual(chartRect.y + chartRect.height + 0.5);

    const dx = 20;
    const dy = 10;

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", { clientX: rect.left + startX, clientY: rect.top + startY, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", { clientX: rect.left + startX + dx, clientY: rect.top + startY + dy, button: 0, pointerId: 1 }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", { clientX: rect.left + startX + dx, clientY: rect.top + startY + dy, button: 0, pointerId: 1 }),
    );

    const persisted = (app.getDocument() as any).getSheetDrawings(sheetId) as any[];
    const after = persisted.find((d) => d?.id === drawing.id) ?? null;
    expect(after).not.toBeNull();
    expect(after.anchor?.type).toBe("absolute");
    const afterAnchor = after.anchor as any;
    const beforeAnchor = drawing.anchor as any;
    const expectedCx = beforeAnchor.size.cx + Math.round(pxToEmu(dx, zoom));
    const expectedCy = beforeAnchor.size.cy + Math.round(pxToEmu(dy, zoom));
    expect(afterAnchor.size.cx).toBeCloseTo(expectedCx, 4);
    expect(afterAnchor.size.cy).toBeCloseTo(expectedCy, 4);

    // Ensure the chart did not move/resize as a side effect of the pointerdown (the handle should win).
    const chartAfter = app.listCharts().find((c) => c.id === chartId) ?? null;
    expect(chartAfter).toBeTruthy();
    expect(chartAfter!.anchor).toEqual(chartAnchorBefore);

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });
});
