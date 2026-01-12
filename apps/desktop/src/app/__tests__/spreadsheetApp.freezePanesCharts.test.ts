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

describe("SpreadsheetApp charts + frozen panes", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
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

  it("keeps chart anchors in frozen panes pinned even if derived frozen counts are stale", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const doc = app.getDocument();
    doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

    (app as any).scrollX = 50;
    (app as any).scrollY = 100;
    (app as any).frozenRows = 0;
    (app as any).frozenCols = 0;

    const rect = (app as any).chartAnchorToViewportRect({
      kind: "twoCell",
      fromCol: 0,
      fromRow: 0,
      fromColOffEmu: 0,
      fromRowOffEmu: 0,
      toCol: 4,
      toRow: 4,
      toColOffEmu: 0,
      toRowOffEmu: 0,
    });

    expect(rect).not.toBeNull();
    expect(rect.left).toBe(0);
    expect(rect.top).toBe(0);

    app.destroy();
    root.remove();
  });

  it("routes charts into pane quadrants in shared-grid mode", () => {
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

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      // Anchor in A1 so it should belong to the top-left frozen pane.
      const result = app.addChart({
        chart_type: "bar",
        data_range: "Sheet1!A2:B5",
        title: "Frozen Chart",
        position: "Sheet1!A1",
      });

      const panes = (app as any).sharedChartPanes as
        | { topLeft: HTMLElement; topRight: HTMLElement; bottomLeft: HTMLElement; bottomRight: HTMLElement }
        | null;
      expect(panes).not.toBeNull();

      const host = ((app as any).chartElements as Map<string, HTMLElement>).get(result.chart_id);
      expect(host).toBeTruthy();
      expect(host?.parentElement).toBe(panes!.topLeft);

      // The outer layer should stay pinned under headers (not under user frozen panes).
      const chartLayer = (app as any).chartLayer as HTMLElement;
      expect(chartLayer.style.left).toBe("48px");
      expect(chartLayer.style.top).toBe("24px");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
