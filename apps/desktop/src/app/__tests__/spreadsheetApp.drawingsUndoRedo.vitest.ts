/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay } from "../../drawings/overlay";
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

describe("SpreadsheetApp drawings undo/redo integration", () => {
  let priorCanvasCharts: string | undefined;
  let priorUseCanvasCharts: string | undefined;

  afterEach(() => {
    if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
    else process.env.CANVAS_CHARTS = priorCanvasCharts;
    if (priorUseCanvasCharts === undefined) delete process.env.USE_CANVAS_CHARTS;
    else process.env.USE_CANVAS_CHARTS = priorUseCanvasCharts;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorCanvasCharts = process.env.CANVAS_CHARTS;
    priorUseCanvasCharts = process.env.USE_CANVAS_CHARTS;
    process.env.CANVAS_CHARTS = "0";
    process.env.USE_CANVAS_CHARTS = "0";
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

  it("stores picture drawings in DocumentController and updates overlay state on undo", async () => {
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
      // Canvas charts are enabled by default, so ChartStore charts appear in the drawings overlay.
      // Remove any charts so the test can focus on picture undo/redo without extra overlay objects.
      for (const chart of app.listCharts()) {
        (app as any).chartStore.deleteChart(chart.id);
      }
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      const overlay = (app as any).drawingOverlay as DrawingOverlay;
      const renderSpy = vi.spyOn(overlay, "render").mockImplementation(() => {});

      const file = new File([new Uint8Array([137, 80, 78, 71])], "test.png", { type: "image/png" });
      await app.insertPicturesFromFiles([file]);

      expect(doc.getSheetDrawings(sheetId)).toHaveLength(1);
      const insertedImages = app.getDrawingObjects(sheetId).filter((obj) => obj.kind.type === "image");
      expect(insertedImages).toHaveLength(1);
      expect(
        renderSpy.mock.calls.some(
          (call) => (call[0] as any[]).filter((obj) => obj?.kind?.type === "image").length === 1,
        ),
      ).toBe(true);

      renderSpy.mockClear();
      doc.undo();

      expect(doc.getSheetDrawings(sheetId)).toHaveLength(0);
      const imagesAfterUndo = app.getDrawingObjects(sheetId).filter((obj) => obj.kind.type === "image");
      expect(imagesAfterUndo).toHaveLength(0);
      expect(
        renderSpy.mock.calls.some(
          (call) => (call[0] as any[]).filter((obj) => obj?.kind?.type === "image").length === 0,
        ),
      ).toBe(true);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
