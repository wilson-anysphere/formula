/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

describe("SpreadsheetApp charts dirty-on-scroll", () => {
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

  it("refreshes a chart's cached data when it was off-screen during edits", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    const sheetToken = /[^A-Za-z0-9_]/.test(sheetId) ? `'${sheetId.replace(/'/g, "''")}'` : sheetId;

    const result = app.addChart({
      chart_type: "bar",
      data_range: `${sheetToken}!A2:B5`,
      title: "Demo",
      position: `${sheetToken}!A1`,
    });

    const chartModels = (app as any).chartModels as Map<string, any>;
    const beforeModel = chartModels.get(result.chart_id);
    expect(beforeModel).toBeTruthy();
    const beforeValues = [...(beforeModel?.series?.[0]?.values?.cache ?? [])];

    // Scroll far enough that the chart is guaranteed to be off-screen.
    (app as any).scrollX = 100_000;
    (app as any).scrollY = 100_000;
    (app as any).renderCharts(false);

    // Mutate a value cell inside the chart's data range (B2).
    doc.setCellValue(sheetId, { row: 1, col: 1 }, 10);
    // Mimic the UI path: a full refresh occurs, but should not rescan chart data while off-screen.
    app.refresh("full");

    const midModel = chartModels.get(result.chart_id);
    expect(midModel).toBeTruthy();
    const midValues = [...(midModel?.series?.[0]?.values?.cache ?? [])];
    expect(midValues).toEqual(beforeValues);

    // Scroll back so the chart becomes visible; the chart should now rescan its data.
    (app as any).scrollX = 0;
    (app as any).scrollY = 0;
    (app as any).renderCharts(false);

    const afterModel = chartModels.get(result.chart_id);
    expect(afterModel).toBeTruthy();
    const afterValues = [...(afterModel?.series?.[0]?.values?.cache ?? [])];
    expect(afterValues).not.toEqual(beforeValues);
    expect(afterValues[0]).toBe(10);

    app.destroy();
    root.remove();
  });
});

