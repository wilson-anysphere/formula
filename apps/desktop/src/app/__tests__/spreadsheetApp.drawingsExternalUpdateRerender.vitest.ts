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

describe("SpreadsheetApp drawings overlay rerender on external sheet view deltas (shared grid)", () => {
  const contexts = new WeakMap<HTMLCanvasElement, any>();

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";

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
      value: function (this: HTMLCanvasElement) {
        const existing = contexts.get(this);
        if (existing) return existing;

        const noop = () => {};
        const gradient = { addColorStop: noop } as any;

        const base: any = {
          canvas: this,
          // Track a few drawing calls so tests can assert rerendering.
          clearRect: vi.fn(),
          strokeRect: vi.fn(),
          drawImage: vi.fn(),
          // Minimal API surface for SpreadsheetApp grid rendering.
          measureText: (text: string) => ({ width: text.length * 8 }),
          createLinearGradient: () => gradient,
          createPattern: () => null,
          getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
          putImageData: noop,
        };

        const ctx = new Proxy(base, {
          get(target, prop) {
            if (prop in target) return target[prop];
            return noop;
          },
          set(target, prop, value) {
            target[prop] = value;
            return true;
          },
        });

        contexts.set(this, ctx);
        return ctx;
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("rerenders drawings when external sheet view deltas arrive (even when refresh() is skipped)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    expect(app.getGridMode()).toBe("shared");

    const drawingsCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--drawings");
    expect(drawingsCanvas).toBeTruthy();
    const drawingsCtx = drawingsCanvas!.getContext("2d") as any;

    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument() as any;
    doc.setSheetDrawings(sheetId, [
      {
        id: "drawing_1",
        zOrder: 0,
        kind: { type: "shape", label: "shape" },
        anchor: { type: "cell", row: 0, col: 0, size: { width: 100, height: 100 } },
      },
    ]);

    const strokeCallsBefore = (drawingsCtx.strokeRect as ReturnType<typeof vi.fn>).mock.calls.length;

    const beforeView = doc.getSheetView(sheetId);
    const afterView = {
      ...beforeView,
      rowHeights: { ...(beforeView?.rowHeights ?? {}), "0": 42 },
    };

    doc.applyExternalSheetViewDeltas([{ sheetId, before: beforeView, after: afterView }], { source: "collab" });

    const strokeCallsAfter = (drawingsCtx.strokeRect as ReturnType<typeof vi.fn>).mock.calls.length;
    expect(strokeCallsAfter).toBeGreaterThan(strokeCallsBefore);

    app.destroy();
    root.remove();
  });
});
