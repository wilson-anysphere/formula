/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";
import { pxToEmu } from "../../drawings/overlay";
import type { DrawingObject } from "../../drawings/types";

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

type CtxCall = { method: string; args: unknown[] };

function createMockCanvasContext(calls: CtxCall[]): CanvasRenderingContext2D {
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
        if (typeof prop === "string") {
          return (...args: unknown[]) => {
            calls.push({ method: prop, args });
          };
        }
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

describe("SpreadsheetApp drawings overlay + shared-grid axis resize", () => {
  const ctxCallsByCanvas = new Map<HTMLCanvasElement, CtxCall[]>();
  const ctxByCanvas = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();

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

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    ctxCallsByCanvas.clear();
    ctxByCanvas.clear();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function getContext(this: HTMLCanvasElement) {
        const existing = ctxByCanvas.get(this);
        if (existing) return existing;
        const calls: CtxCall[] = [];
        const ctx = createMockCanvasContext(calls);
        ctxCallsByCanvas.set(this, calls);
        ctxByCanvas.set(this, ctx);
        return ctx;
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("re-renders drawings when shared-grid column widths change", async () => {
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

      const drawingCanvas = (app as any).drawingCanvas as HTMLCanvasElement;
      expect(drawingCanvas).toBeTruthy();

      const objects: DrawingObject[] = [
        {
          id: 1,
          kind: { type: "image", imageId: "missing" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 1 }, offset: { xEmu: 0, yEmu: 0 } }, // B1
            size: { cx: pxToEmu(10), cy: pxToEmu(10) },
          },
          zOrder: 0,
        },
      ];

      // Seed the drawing anchored at B1 via a monkeypatched drawings getter. The production API
      // (`DocumentController.getSheetDrawings`) is not yet stable, so SpreadsheetApp treats it
      // as an optional integration point.
      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();
      calls!.splice(0, calls!.length);

      const renderSpy = vi.spyOn(app as any, "renderDrawings");

      // Initial render.
      (app as any).renderDrawings();
      // DrawingOverlay may await async image hydration (IndexedDB) before emitting placeholder
      // stroke calls; yield to the event loop so any pending microtasks complete.
      await new Promise((resolve) => setTimeout(resolve, 0));
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const x1 = Number(firstStroke!.args[0]);
      expect(Number.isFinite(x1)).toBe(true);

      // Simulate an interactive drag that resized column A (doc col 0 => grid col 1).
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1;
      const prevSize = renderer.getColWidth(index);
      const nextSize = prevSize + 50;
      renderer.setColWidth(index, nextSize);

      // Clear previous calls and spy counts.
      calls!.splice(0, calls!.length);
      renderSpy.mockClear();

      (app as any).onSharedGridAxisSizeChange({
        kind: "col",
        index,
        size: nextSize,
        previousSize: prevSize,
        defaultSize: renderer.scroll.cols.defaultSize,
        zoom: renderer.getZoom(),
        source: "resize",
      });
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(renderSpy).toHaveBeenCalled();
      const secondStroke = calls!.find((call) => call.method === "strokeRect");
      expect(secondStroke).toBeTruthy();
      const x2 = Number(secondStroke!.args[0]);
      expect(x2).toBeCloseTo(x1 + (nextSize - prevSize), 6);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
