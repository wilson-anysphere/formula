/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../../grid/splitView/secondaryGridView.js";
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

describe("SpreadsheetApp.hitTestDrawingAtClientPoint (split view)", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

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
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("treats the split-view rotation handle as a hit target even in legacy mode", () => {
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
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(drawing.id);

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
    });

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

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
    } as any;

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
        return { width: view.grid.renderer.getColWidth(gridCol), height: view.grid.renderer.getRowHeight(gridRow) };
      },
    };

    const bounds = drawingObjectToViewportRect(drawing, viewport, geom as any);
    const handleCenter = getRotationHandleCenter(bounds, drawing.transform);
    expect(Number.isFinite(handleCenter.x)).toBe(true);
    expect(Number.isFinite(handleCenter.y)).toBe(true);

    const hit = app.hitTestDrawingAtClientPoint(rect.left + handleCenter.x, rect.top + handleCenter.y);
    expect(hit).toEqual({ id: drawing.id });

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });

  it("clears the selected drawing via Escape key from the split-view secondary pane", () => {
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
    app.setDrawingObjects([drawing]);
    app.selectDrawingById(drawing.id);
    expect(app.getSelectedDrawingId()).toBe(drawing.id);

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
    });

    const selectionCanvas = secondaryContainer.querySelector<HTMLCanvasElement>("canvas.grid-canvas--selection");
    if (!selectionCanvas) throw new Error("Missing secondary selection canvas");

    app.setSplitViewSecondaryGridView({ container: secondaryContainer, grid: view.grid });

    const esc = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    selectionCanvas.dispatchEvent(esc);
    expect(esc.defaultPrevented).toBe(true);
    expect(app.getSelectedDrawingId()).toBeNull();

    view.destroy();
    app.destroy();
    secondaryContainer.remove();
    root.remove();
  });
});
