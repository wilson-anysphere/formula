/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

describe("SpreadsheetApp (canvas charts) formats in-cell images for chart data reads", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";
    process.env.CANVAS_CHARTS = "1";

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

  it("uses image altText instead of `[object Object]` in category caches", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const doc: any = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, "Category");
    doc.setCellValue(sheetId, { row: 0, col: 1 }, "Value");
    doc.setCellValue(sheetId, { row: 1, col: 0 }, { type: "image", value: { imageId: "img_1", altText: "Kitten" } });
    doc.setCellValue(sheetId, { row: 1, col: 1 }, 1);
    doc.setCellValue(sheetId, { row: 2, col: 0 }, "Dog");
    doc.setCellValue(sheetId, { row: 2, col: 1 }, 2);

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A1:B3", title: "Image Chart" });

    const model = (app as any).chartCanvasStoreAdapter.getChartModel(chartId) as any;
    expect(model?.series?.[0]?.categories?.cache).toEqual(["Kitten", "Dog"]);

    app.destroy();
    root.remove();
  });

  it("falls back to cached image payloads for formula cells when computed values are blank", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const doc: any = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, "Category");
    doc.setCellValue(sheetId, { row: 0, col: 1 }, "Value");
    // Simulate a formula cell with a cached IMAGE() payload (formula result is blank/unsupported).
    doc.setCell(sheetId, 1, 0, { formula: '=""', value: { type: "image", value: { imageId: "img_1", altText: "Kitten" } } });
    doc.setCellValue(sheetId, { row: 1, col: 1 }, 1);
    doc.setCellValue(sheetId, { row: 2, col: 0 }, "Dog");
    doc.setCellValue(sheetId, { row: 2, col: 1 }, 2);

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A1:B3", title: "Image Chart" });

    const model = (app as any).chartCanvasStoreAdapter.getChartModel(chartId) as any;
    expect(model?.series?.[0]?.categories?.cache).toEqual(["Kitten", "Dog"]);

    app.destroy();
    root.remove();
  });
});

