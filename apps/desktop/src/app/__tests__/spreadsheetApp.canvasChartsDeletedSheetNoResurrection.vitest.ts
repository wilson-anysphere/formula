/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { createSheetNameResolverFromIdToNameMap } from "../../sheet/sheetNameResolver";
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

function createStatus() {
  return {
    activeCell: document.createElement("div"),
    selectionRange: document.createElement("div"),
    activeValue: document.createElement("div"),
  };
}

describe("SpreadsheetApp canvas charts deleted-sheet safety", () => {
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

  it("does not recreate a deleted sheet when charts refresh using a stale sheetNameResolver mapping", () => {
    // Canvas charts are enabled by default, but set the env var explicitly so this test remains
    // robust even if other suites temporarily override the default.
    process.env.CANVAS_CHARTS = "1";

    const staleIdToName = new Map<string, string>([
      ["Sheet1", "Sheet1"],
      ["Sheet2", "Sheet2"],
    ]);
    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(staleIdToName);

    const app = new SpreadsheetApp(createRoot(), createStatus(), { sheetNameResolver });
    try {
      expect((app as any).useCanvasCharts).toBe(true);
      const doc = app.getDocument();

      // Ensure the default sheet is materialized (DocumentController creates sheets lazily).
      doc.getCell("Sheet1", { row: 0, col: 0 });
      doc.addSheet({ sheetId: "Sheet2", name: "Sheet2", insertAfterId: "Sheet1" });

      doc.setCellValue("Sheet2", { row: 0, col: 0 }, "A");
      doc.setCellValue("Sheet2", { row: 0, col: 1 }, 1);
      doc.setCellValue("Sheet2", { row: 1, col: 0 }, "B");
      doc.setCellValue("Sheet2", { row: 1, col: 1 }, 2);

      const { chart_id: chartId } = app.addChart({
        chart_type: "bar",
        data_range: "Sheet2!A1:B2",
        position: "Sheet1!C1",
        title: "Deleted Sheet Data Chart",
      });

      const adapter = (app as any).chartCanvasStoreAdapter;
      adapter.getChartModel(chartId);

      doc.deleteSheet("Sheet2");
      expect(doc.getSheetIds()).toEqual(["Sheet1"]);
      expect(doc.getSheetMeta("Sheet2")).toBeNull();

      // Force a chart content rebuild (the adapter will consult the stale resolver and attempt to
      // read values from Sheet2). This must NOT resurrect the deleted sheet.
      adapter.invalidate(chartId);
      adapter.getChartModel(chartId);

      expect(doc.getSheetIds()).toEqual(["Sheet1"]);
      expect(doc.getSheetMeta("Sheet2")).toBeNull();
    } finally {
      app.destroy();
    }
  });
});
