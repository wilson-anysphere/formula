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

describe("SpreadsheetApp collab repaint", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: vi.fn((cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      }),
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

  it("schedules a repaint when externally-sourced collab deltas arrive (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");
    const rafSpy = globalThis.requestAnimationFrame as unknown as ReturnType<typeof vi.fn>;

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    const before = doc.getCell(sheetId, { row: 0, col: 1 }) as any;

    doc.applyExternalDeltas(
      [
        {
          sheetId,
          row: 0,
          col: 1,
          before,
          after: { value: "Remote", formula: null, styleId: before?.styleId ?? 0 },
        },
      ],
      { source: "collab" },
    );

    expect(refreshSpy).toHaveBeenCalledWith("scroll");
    expect(rafSpy).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("schedules a repaint when pivot-driven updates arrive (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, "Pivot Result", { source: "pivot" });

    expect(refreshSpy).toHaveBeenCalledWith("scroll");

    app.destroy();
    root.remove();
  });

  it("schedules a repaint when extension-driven updates arrive (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();

    doc.setCellValue(sheetId, { row: 0, col: 0 }, "From Extension", { source: "extension" });

    expect(refreshSpy).toHaveBeenCalledWith("scroll");

    app.destroy();
    root.remove();
  });

  it("schedules a repaint when sheet-rename formula rewrites arrive (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument() as any;
    const sheetId = app.getCurrentSheetId();

    doc.setCellInputs(
      [{ sheetId, row: 0, col: 0, value: "Renamed Sheet Rewrite", formula: null }],
      { label: "Rename Sheet", source: "sheetRename" },
    );

    expect(refreshSpy).toHaveBeenCalledWith("scroll");

    app.destroy();
    root.remove();
  });

  it("schedules a repaint when sheet-delete formula rewrites arrive (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument() as any;
    const sheetId = app.getCurrentSheetId();

    doc.setCellInputs(
      [{ sheetId, row: 0, col: 0, value: "Deleted Sheet Rewrite", formula: null }],
      { label: "Delete Sheet", source: "sheetDelete" },
    );

    expect(refreshSpy).toHaveBeenCalledWith("scroll");

    app.destroy();
    root.remove();
  });

  it("schedules a repaint when applyState replaces workbook contents (legacy grid)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument() as any;
    const snapshot = doc.encodeState() as Uint8Array;

    const decoded = new TextDecoder().decode(snapshot);
    const parsed = JSON.parse(decoded) as any;
    const sheetId = app.getCurrentSheetId();
    const sheet = (parsed?.sheets ?? []).find((s: any) => s && s.id === sheetId);
    expect(sheet).toBeTruthy();

    const cells: any[] = Array.isArray(sheet.cells) ? sheet.cells : [];
    const target = cells.find((c) => c && c.row === 0 && c.col === 0);
    expect(target).toBeTruthy();
    target.value = "From ApplyState";
    target.formula = null;

    const encoded = new TextEncoder().encode(JSON.stringify(parsed));
    doc.applyState(encoded);

    expect(refreshSpy).toHaveBeenCalledWith("scroll");

    app.destroy();
    root.remove();
  });

  it("keeps the status bar in sync when a remote edit changes an active cell dependency", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const refreshSpy = vi.spyOn(app, "refresh");

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    // Make the active cell depend on B2.
    doc.setCellFormula(sheetId, { row: 0, col: 0 }, "=B2");
    app.refresh();
    expect(status.activeValue.textContent).toBe("2");

    const before = doc.getCell(sheetId, { row: 1, col: 1 }) as any;

    doc.applyExternalDeltas(
      [
        {
          sheetId,
          row: 1,
          col: 1,
          before,
          after: { value: 99, formula: null, styleId: before?.styleId ?? 0 },
        },
      ],
      { source: "collab" },
    );

    expect(refreshSpy).toHaveBeenCalledWith("scroll");
    expect(status.activeValue.textContent).toBe("99");

    app.destroy();
    root.remove();
  });

  it("re-renders chart SVG content when remote edits change chart data", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("legacy");

    const { chart_id: chartId } = app.addChart({ chart_type: "bar", data_range: "A2:B5", title: "Test Chart" });

    const beforeModel = getChartModel(app, chartId);
    expect(beforeModel).toBeTruthy();
    const beforeValues = [...(beforeModel?.series?.[0]?.values?.cache ?? [])];

    const doc = app.getDocument();
    const sheetId = app.getCurrentSheetId();
    // Mutate one of the chart's value cells (B2 is part of Sheet1!A2:B5).
    const before = doc.getCell(sheetId, { row: 1, col: 1 }) as any;

    doc.applyExternalDeltas(
      [
        {
          sheetId,
          row: 1,
          col: 1,
          before,
          after: { value: 10, formula: null, styleId: before?.styleId ?? 0 },
        },
      ],
      { source: "collab" },
    );

    const afterModel = getChartModel(app, chartId);
    expect(afterModel).toBeTruthy();
    const afterValues = [...(afterModel?.series?.[0]?.values?.cache ?? [])];
    expect(afterValues).not.toEqual(beforeValues);
    expect(afterValues[0]).toBe(10);

    app.destroy();
    root.remove();
  });

  it("does not call SpreadsheetApp.refresh for external deltas in shared-grid mode", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      const refreshSpy = vi.spyOn(app, "refresh");

      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const before = doc.getCell(sheetId, { row: 0, col: 0 }) as any;

      doc.applyExternalDeltas(
        [
          {
            sheetId,
            row: 0,
            col: 0,
            before,
            after: { value: "Remote Shared", formula: null, styleId: before?.styleId ?? 0 },
          },
        ],
        { source: "collab" },
      );

      expect(refreshSpy).not.toHaveBeenCalled();

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
