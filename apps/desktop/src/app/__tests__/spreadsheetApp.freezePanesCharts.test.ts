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

type CanvasCall = { method: string; args: unknown[] };

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

function createRecordingCanvasContext(calls: CanvasCall[]): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const record =
    (method: string) =>
    (...args: unknown[]) => {
      calls.push({ method, args });
    };

  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      calls,
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
      save: record("save"),
      restore: record("restore"),
      beginPath: record("beginPath"),
      rect: record("rect"),
      clip: record("clip"),
      translate: record("translate"),
      clearRect: record("clearRect"),
      setTransform: record("setTransform"),
      scale: record("scale"),
      drawImage: record("drawImage"),
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

    const chartCalls: CanvasCall[] = [];
    (globalThis as any).__chartCanvasCalls = chartCalls;
    const chartContexts = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      value: function (this: HTMLCanvasElement): any {
        if (this.className.includes("grid-canvas--chart")) {
          const existing = chartContexts.get(this);
          if (existing) return existing;
          const created = createRecordingCanvasContext(chartCalls);
          chartContexts.set(this, created);
          return created;
        }
        return createMockCanvasContext();
      },
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

      const renderSpy = vi.spyOn((app as any).chartRenderer, "renderToCanvas");
      renderSpy.mockClear();
      (globalThis as any).__chartCanvasCalls.length = 0;

      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      // Anchor in A1 so it should belong to the top-left frozen pane.
      const result = app.addChart({
        chart_type: "bar",
        data_range: "Sheet1!A2:B5",
        title: "Frozen Chart",
        position: "Sheet1!A1",
      });

      // Canvas-only: no DOM chart hosts should exist.
      expect(root.querySelector('[data-testid="chart-object"]')).toBeNull();

      // In shared-grid mode, charts are rendered into the chart canvas with a translation
      // that pins the overlay under the frozen header row/column (48px x 24px).
      const calls: CanvasCall[] = (globalThis as any).__chartCanvasCalls;
      expect(calls.some((c) => c.method === "translate" && c.args[0] === 48 && c.args[1] === 24)).toBe(true);

      // The chart anchored in A1 should be routed to the top-left frozen quadrant, clipped
      // to the first frozen column/row (1 col x 1 row => 100px x 24px in the data area).
      expect(calls.some((c) => c.method === "rect" && c.args[0] === 0 && c.args[1] === 0 && c.args[2] === 100 && c.args[3] === 24)).toBe(true);

      const chartCall = renderSpy.mock.calls.find((args) => args[1] === result.chart_id);
      expect(chartCall).toBeTruthy();
      expect(chartCall?.[2]).toMatchObject({ x: 0, y: 0 });

      const selectionCanvas = (app as any).selectionCanvas as HTMLElement;
      expect(selectionCanvas.classList.contains("grid-canvas--shared-selection")).toBe(true);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
