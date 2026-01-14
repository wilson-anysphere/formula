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

describe("SpreadsheetApp chart keyboard shortcuts (canvas charts)", () => {
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

  it("duplicates the selected chart on Ctrl+D (overrides fill down)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Duplicate Chart",
      position: "C1",
    });

    app.selectDrawingById(chartIdToDrawingId(chartId));
    expect(app.getSelectedChartId()).toBe(chartId);

    const activeBefore = app.getActiveCell();
    const countBefore = app.listCharts().length;
    const anchorBefore = JSON.parse(JSON.stringify(app.listCharts().find((c) => c.id === chartId)?.anchor ?? null));

    root.dispatchEvent(new KeyboardEvent("keydown", { key: "d", ctrlKey: true, bubbles: true, cancelable: true }));

    expect(app.listCharts().length).toBe(countBefore + 1);
    const selectedAfter = app.getSelectedChartId();
    expect(selectedAfter).toBeTruthy();
    expect(selectedAfter).not.toBe(chartId);

    const duplicated = app.listCharts().find((c) => c.id === selectedAfter);
    expect(duplicated).toBeTruthy();
    expect(JSON.parse(JSON.stringify(duplicated!.anchor))).not.toEqual(anchorBefore);
    expect(app.getActiveCell()).toEqual(activeBefore);

    app.destroy();
    root.remove();
  });

  it("uses Ctrl+BracketRight to bring charts forward (overrides auditing)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    const { chart_id: chartA } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart A",
      position: "C1",
    });
    const { chart_id: chartB } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Chart B",
      position: "C15",
    });

    app.selectDrawingById(chartIdToDrawingId(chartA));
    expect(app.getSelectedChartId()).toBe(chartA);

    const orderBefore = app
      .listCharts()
      .filter((c) => c.sheetId === sheetId)
      .map((c) => c.id);

    expect(orderBefore.indexOf(chartA)).toBeGreaterThanOrEqual(0);
    expect(orderBefore.indexOf(chartB)).toBeGreaterThanOrEqual(0);
    expect(orderBefore.indexOf(chartA)).toBeLessThan(orderBefore.indexOf(chartB));

    root.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "]",
        code: "BracketRight",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      }),
    );

    const orderAfter = app
      .listCharts()
      .filter((c) => c.sheetId === sheetId)
      .map((c) => c.id);

    expect(orderAfter.indexOf(chartA)).toBeGreaterThan(orderAfter.indexOf(chartB));
    expect(app.getSelectedChartId()).toBe(chartA);

    app.destroy();
    root.remove();
  });
});

