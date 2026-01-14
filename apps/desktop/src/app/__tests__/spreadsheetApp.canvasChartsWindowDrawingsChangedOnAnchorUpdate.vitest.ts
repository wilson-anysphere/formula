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

describe("SpreadsheetApp (canvas charts) drawings-changed window event", () => {
  let rafCallbacks: FrameRequestCallback[] = [];
  const flushRaf = () => {
    let iterations = 0;
    while (rafCallbacks.length > 0) {
      if (iterations++ > 50) throw new Error("flushRaf: too many iterations");
      const cbs = rafCallbacks;
      rafCallbacks = [];
      for (const cb of cbs) cb(0);
    }
  };

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

    rafCallbacks = [];
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      writable: true,
      value: (cb: FrameRequestCallback) => {
        rafCallbacks.push(cb);
        return rafCallbacks.length;
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

  it("dispatches formula:drawings-changed on canvas chart anchor updates (coalesced)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    flushRaf();

    const { chart_id: chartId } = app.addChart({
      chart_type: "bar",
      data_range: "A2:B5",
      title: "Window Event Chart",
      position: "A1",
    });
    flushRaf();

    const onChanged = vi.fn();
    window.addEventListener("formula:drawings-changed", onChanged);

    try {
      const store = (app as any).chartStore as any;
      store.updateChartAnchor(chartId, { kind: "absolute", xEmu: 1, yEmu: 2, cxEmu: 3, cyEmu: 4 });
      store.updateChartAnchor(chartId, { kind: "absolute", xEmu: 5, yEmu: 6, cxEmu: 7, cyEmu: 8 });
      expect(onChanged).toHaveBeenCalledTimes(0);

      flushRaf();
      expect(onChanged).toHaveBeenCalledTimes(1);

      store.updateChartAnchor(chartId, { kind: "absolute", xEmu: 9, yEmu: 10, cxEmu: 11, cyEmu: 12 });
      flushRaf();
      expect(onChanged).toHaveBeenCalledTimes(2);
    } finally {
      window.removeEventListener("formula:drawings-changed", onChanged);
      app.destroy();
      root.remove();
    }
  });
});

