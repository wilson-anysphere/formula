/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { chartIdToDrawingId } from "../../charts/chartDrawingAdapter";
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

describe("SpreadsheetApp canvas charts (default)", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    // Keep the grid mode deterministic for test geometry expectations.
    process.env.DESKTOP_GRID_MODE = "legacy";
    delete process.env.CANVAS_CHARTS;
    delete process.env.USE_CANVAS_CHARTS;

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

  it("defaults to canvas charts and renders ChartStore charts via the drawings overlay", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect((app as any).useCanvasCharts).toBe(true);

    // Legacy chart canvases should not be mounted by default.
    expect(root.querySelector(".grid-canvas--chart")).toBeNull();
    expect(root.querySelector(".chart-selection-canvas")).toBeNull();

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });
    const chart = app.listCharts().find((c) => c.id === chartId) ?? null;
    expect(chart).not.toBeNull();

    const drawingObjects = app.getDrawingObjects();
    const chartDrawingId = chartIdToDrawingId(chartId);
    const chartObj = drawingObjects.find((obj) => obj.kind.type === "chart" && obj.id === chartDrawingId) ?? null;
    expect(chartObj).not.toBeNull();

    // Selection should still work through the unified drawing selection stack.
    app.selectDrawingById(chartDrawingId);
    expect(app.getSelectedChartId()).toBe(chartId);
    expect(app.getDrawingHandlePointsPx(chartDrawingId)).not.toBeNull();

    app.destroy();
    root.remove();
  });
});
