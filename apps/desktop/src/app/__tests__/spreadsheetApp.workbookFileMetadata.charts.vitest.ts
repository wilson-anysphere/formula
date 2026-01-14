/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

function getChartModel(app: SpreadsheetApp, chartId: string): any {
  const anyApp = app as any;
  if (anyApp.useCanvasCharts) {
    return anyApp.chartCanvasStoreAdapter.getChartModel(chartId);
  }
  return (anyApp.chartModels as Map<string, any>).get(chartId);
}

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

describe("SpreadsheetApp chart refresh on workbook metadata changes", () => {
  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;

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

  it("refreshes a visible chart's cached data when the workbook file metadata changes (legacy grid)", async () => {
    process.env.DESKTOP_GRID_MODE = "legacy";

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    const sheetToken = /[^A-Za-z0-9_]/.test(sheetId) ? `'${sheetId.replace(/'/g, "''")}'` : sheetId;

    doc.setCellFormula(sheetId, { row: 0, col: 0 }, '=CELL("filename")');
    doc.setCellValue(sheetId, { row: 0, col: 1 }, 1);

    const chart = app.addChart({
      chart_type: "bar",
      data_range: `${sheetToken}!A1:B1`,
      title: "Workbook Metadata Chart",
      position: `${sheetToken}!C1`,
    });

    expect(getChartModel(app, chart.chart_id)?.series?.[0]?.categories?.cache?.[0]).toBe("");

    await app.setWorkbookFileMetadata("/tmp", "Book.xlsx");
    const sheetName = app.getCurrentSheetDisplayName();
    expect(getChartModel(app, chart.chart_id)?.series?.[0]?.categories?.cache?.[0]).toBe(`/tmp/[Book.xlsx]${sheetName}`);

    app.destroy();
    root.remove();
  });

  it("refreshes a visible chart's cached data when the workbook file metadata changes (shared grid)", async () => {
    process.env.DESKTOP_GRID_MODE = "shared";

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");
    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    const sheetToken = /[^A-Za-z0-9_]/.test(sheetId) ? `'${sheetId.replace(/'/g, "''")}'` : sheetId;

    doc.setCellFormula(sheetId, { row: 0, col: 0 }, '=CELL("filename")');
    doc.setCellValue(sheetId, { row: 0, col: 1 }, 1);

    const chart = app.addChart({
      chart_type: "bar",
      data_range: `${sheetToken}!A1:B1`,
      title: "Workbook Metadata Chart",
      position: `${sheetToken}!C1`,
    });

    expect(getChartModel(app, chart.chart_id)?.series?.[0]?.categories?.cache?.[0]).toBe("");

    await app.setWorkbookFileMetadata("/tmp", "Book.xlsx");
    const sheetName = app.getCurrentSheetDisplayName();
    expect(getChartModel(app, chart.chart_id)?.series?.[0]?.categories?.cache?.[0]).toBe(`/tmp/[Book.xlsx]${sheetName}`);

    app.destroy();
    root.remove();
  });
});
