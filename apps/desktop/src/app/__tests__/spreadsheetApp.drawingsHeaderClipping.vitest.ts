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

type RecordedCall = { method: string; args: unknown[] };

function createRecordingCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const calls: RecordedCall[] = [];
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;

  const record =
    (method: string) =>
    (...args: unknown[]) => {
      calls.push({ method, args });
    };

  const target: any = {
    canvas,
    __calls: calls,
    save: record("save"),
    restore: record("restore"),
    beginPath: record("beginPath"),
    rect: record("rect"),
    clip: record("clip"),
    clearRect: record("clearRect"),
    strokeRect: record("strokeRect"),
    fillText: record("fillText"),
    setLineDash: record("setLineDash"),
    drawImage: record("drawImage"),
    measureText: (text: string) => ({ width: text.length * 8 }),
    createLinearGradient: () => gradient,
    createPattern: () => null,
    getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
    putImageData: noop,
  };

  return new Proxy(target, {
    get(obj, prop) {
      if (prop in obj) return obj[prop];
      return noop;
    },
    set(obj, prop, value) {
      obj[prop] = value;
      return true;
    },
  }) as any;
}

function createRoot(rect: { width: number; height: number } = { width: 800, height: 600 }): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: rect.width,
      height: rect.height,
      left: 0,
      top: 0,
      right: rect.width,
      bottom: rect.height,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

function seedTopLeftDrawing(): DrawingObject {
  return {
    id: 1,
    zOrder: 0,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(20), cy: pxToEmu(20) },
    },
  };
}

describe("SpreadsheetApp drawings header clipping", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.head.innerHTML = "";
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

    // jsdom (as used by vitest) does not provide PointerEvent in all environments.
    // SpreadsheetApp only relies on MouseEvent fields (clientX/Y, button) for drawing hit tests.
    if (typeof (globalThis as any).PointerEvent === "undefined") {
      Object.defineProperty(globalThis, "PointerEvent", { configurable: true, value: MouseEvent });
    }

    // Reuse a stable 2d context per canvas so SpreadsheetApp and DrawingOverlay share call logs.
    const ctxCache = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value(this: HTMLCanvasElement, type: string) {
        if (type !== "2d") return null;
        let ctx = ctxCache.get(this);
        if (!ctx) {
          ctx = createRecordingCanvasContext(this);
          ctxCache.set(this, ctx);
        }
        return ctx;
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("clips drawings to the cell body and ignores header hit tests (legacy grid)", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("legacy");

      const doc: any = app.getDocument();
      doc.getSheetDrawings = (sheetId: string) => (sheetId === app.getCurrentSheetId() ? [seedTopLeftDrawing()] : []);

      app.setScroll(10, 0);

      const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]')!;
      const ctx = drawingCanvas.getContext("2d") as any;
      const calls: RecordedCall[] = ctx.__calls;
      calls.length = 0;

      // Force a single deterministic render pass with the current state.
      (app as any).renderDrawings();

      const headerWidth = (app as any).rowHeaderWidth as number;
      const headerHeight = (app as any).colHeaderHeight as number;
      const rootRect = root.getBoundingClientRect();

      const clipRect = calls.find(
        (c) =>
          c.method === "rect" &&
          c.args[0] === headerWidth &&
          c.args[1] === headerHeight &&
          c.args[2] === rootRect.width - headerWidth &&
          c.args[3] === rootRect.height - headerHeight,
      );
      expect(clipRect).toBeTruthy();

      const rectIdx = calls.indexOf(clipRect!);
      const clipIdx = calls.findIndex((c, idx) => idx > rectIdx && c.method === "clip");
      const strokeIdx = calls.findIndex((c, idx) => idx > clipIdx && c.method === "strokeRect");
      expect(clipIdx).toBeGreaterThan(rectIdx);
      expect(strokeIdx).toBeGreaterThan(clipIdx);

      // Header hit-test should ignore drawings.
      expect((app as any).selectedDrawingId).toBe(null);
      // Legacy grid mode handles drawing selection via the bubbling `onPointerDown` handler.
      (app as any).onPointerDown(
        new PointerEvent("pointerdown", {
          bubbles: true,
          cancelable: true,
          button: 0,
          clientX: headerWidth - 5,
          clientY: headerHeight + 5,
        }),
      );
      expect((app as any).selectedDrawingId).toBe(null);

      // Body hit-test should still select it.
      (app as any).onPointerDown(
        new PointerEvent("pointerdown", {
          bubbles: true,
          cancelable: true,
          button: 0,
          clientX: headerWidth + 5,
          clientY: headerHeight + 5,
        }),
      );
      expect((app as any).selectedDrawingId).toBe(1);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("clips drawings to the cell body and ignores header hit tests (shared grid)", () => {
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

      const doc: any = app.getDocument();
      doc.getSheetDrawings = (sheetId: string) => (sheetId === app.getCurrentSheetId() ? [seedTopLeftDrawing()] : []);

      app.setScroll(10, 0);

      const sharedGrid = (app as any).sharedGrid;
      const sharedViewport = sharedGrid.renderer.scroll.getViewportState();
      const headerWidth = sharedGrid.renderer.getColWidth(0);
      const headerHeight = sharedGrid.renderer.getRowHeight(0);

      const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]')!;
      const ctx = drawingCanvas.getContext("2d") as any;
      const calls: RecordedCall[] = ctx.__calls;
      calls.length = 0;

      (app as any).renderDrawings(sharedViewport);

      const clipRect = calls.find(
        (c) =>
          c.method === "rect" &&
          c.args[0] === headerWidth &&
          c.args[1] === headerHeight &&
          c.args[2] === sharedViewport.width - headerWidth &&
          c.args[3] === sharedViewport.height - headerHeight,
      );
      expect(clipRect).toBeTruthy();

      const rectIdx = calls.indexOf(clipRect!);
      const clipIdx = calls.findIndex((c, idx) => idx > rectIdx && c.method === "clip");
      const strokeIdx = calls.findIndex((c, idx) => idx > clipIdx && c.method === "strokeRect");
      expect(clipIdx).toBeGreaterThan(rectIdx);
      expect(strokeIdx).toBeGreaterThan(clipIdx);

      expect((app as any).selectedDrawingId).toBe(null);
      (app as any).onDrawingPointerDownCapture(
        new PointerEvent("pointerdown", {
          bubbles: true,
          cancelable: true,
          button: 0,
          clientX: headerWidth - 5,
          clientY: headerHeight + 5,
        }),
      );
      expect((app as any).selectedDrawingId).toBe(null);

      (app as any).onDrawingPointerDownCapture(
        new PointerEvent("pointerdown", {
          bubbles: true,
          cancelable: true,
          button: 0,
          clientX: headerWidth + 5,
          clientY: headerHeight + 5,
        }),
      );
      expect((app as any).selectedDrawingId).toBe(1);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
