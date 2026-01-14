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

      // DocumentImageStore supports async hydration (IndexedDB) via `getAsync`, which can
      // schedule follow-up overlay renders after the initial synchronous placeholder pass.
      // Disable it so this unit test can assert deterministically on canvas calls.
      (app as any).drawingImages.getAsync = undefined;

      const drawingCanvas = (app as any).drawingCanvas as HTMLCanvasElement;
      expect(drawingCanvas).toBeTruthy();

      const objects: DrawingObject[] = [
        {
          id: 1,
          kind: { type: "image", imageId: "missing" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 0, col: 1 }, offset: { xEmu: 0, yEmu: 0 } }, // B1
            // Ensure the object has a meaningful interior so cursor hit testing doesn't
            // accidentally land on a resize handle when sampling points within the rect.
            size: { cx: pxToEmu(40), cy: pxToEmu(40) },
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

      const waitForStrokeRect = async (): Promise<CtxCall | undefined> => {
        // Drawing overlay rendering is synchronous, but missing images may trigger async hydration + a
        // subsequent repaint. In shared-grid mode, additional viewport callbacks can also trigger
        // redraws that abort in-flight renders. Poll across a few microtasks for a completed pass.
        for (let i = 0; i < 8; i += 1) {
          const stroke = calls!.find((call) => call.method === "strokeRect");
          if (stroke) return stroke;
          await Promise.resolve();
        }
        return undefined;
      };

      // Initial render.
      (app as any).renderDrawings();
      const firstStroke = await waitForStrokeRect();
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

      expect(renderSpy).toHaveBeenCalled();
      const secondStroke = await waitForStrokeRect();
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

  it("updates drawings during interactive shared-grid column resize (viewport changes)", async () => {
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

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();

      // Baseline render (establish initial x position).
      calls!.splice(0, calls!.length);
      (app as any).renderDrawings();
      await new Promise((resolve) => setTimeout(resolve, 0));
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const x1 = Number(firstStroke!.args[0]);
      expect(Number.isFinite(x1)).toBe(true);

      // Simulate interactive resize drag: update the renderer widths directly.
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1; // grid col 1 => doc col 0 (A)
      const prevSize = renderer.getColWidth(index);
      const nextSize = prevSize + 50;

      calls!.splice(0, calls!.length);
      renderer.setColWidth(index, nextSize);
      await new Promise((resolve) => setTimeout(resolve, 0));

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

  it("updates drawings during interactive shared-grid row resize (viewport changes)", async () => {
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

      // Anchor at A2 so its y-position depends on the height of row 1.
      const objects: DrawingObject[] = [
        {
          id: 1,
          kind: { type: "image", imageId: "missing" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 1, col: 0 }, offset: { xEmu: 0, yEmu: 0 } }, // A2
            size: { cx: pxToEmu(10), cy: pxToEmu(10) },
          },
          zOrder: 0,
        },
      ];

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();

      // Baseline render (establish initial y position).
      calls!.splice(0, calls!.length);
      (app as any).renderDrawings();
      await new Promise((resolve) => setTimeout(resolve, 0));
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const y1 = Number(firstStroke!.args[1]);
      expect(Number.isFinite(y1)).toBe(true);

      // Simulate interactive resize drag: update the renderer heights directly.
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1; // grid row 1 => doc row 0 (row 1)
      const prevSize = renderer.getRowHeight(index);
      const nextSize = prevSize + 30;

      calls!.splice(0, calls!.length);
      renderer.setRowHeight(index, nextSize);
      await new Promise((resolve) => setTimeout(resolve, 0));

      const secondStroke = calls!.find((call) => call.method === "strokeRect");
      expect(secondStroke).toBeTruthy();
      const y2 = Number(secondStroke!.args[1]);
      expect(y2).toBeCloseTo(y1 + (nextSize - prevSize), 6);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("updates drawing hit-testing geometry when shared-grid column widths change", async () => {
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

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();
      calls!.splice(0, calls!.length);

      // Baseline render.
      (app as any).renderDrawings();
      await new Promise((resolve) => setTimeout(resolve, 0));
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const x1 = Number(firstStroke!.args[0]);
      const y1 = Number(firstStroke!.args[1]);
      const w1 = Number(firstStroke!.args[2]);
      const h1 = Number(firstStroke!.args[3]);
      expect(Number.isFinite(x1)).toBe(true);
      expect(Number.isFinite(y1)).toBe(true);
      expect(Number.isFinite(w1)).toBe(true);
      expect(Number.isFinite(h1)).toBe(true);
      // Cursor should detect the drawing at its current position.
      // Use the center of the rect to avoid hitting resize handles (the placeholder is only 10x10px).
      const hitX1 = x1 + w1 / 2;
      const hitY1 = y1 + h1 / 2;
      expect((app as any).drawingCursorAtPoint(hitX1, hitY1)).toBe("move");

      // Resize column A (doc col 0 => grid col 1).
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1;
      const prevSize = renderer.getColWidth(index);
      const nextSize = prevSize + 50;
      renderer.setColWidth(index, nextSize);

      calls!.splice(0, calls!.length);

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

      const secondStroke = calls!.find((call) => call.method === "strokeRect");
      expect(secondStroke).toBeTruthy();
      const x2 = Number(secondStroke!.args[0]);
      const y2 = Number(secondStroke!.args[1]);
      const w2 = Number(secondStroke!.args[2]);
      const h2 = Number(secondStroke!.args[3]);
      expect(x2).toBeCloseTo(x1 + (nextSize - prevSize), 6);
      expect(y2).toBeCloseTo(y1, 6);
      expect(w2).toBeCloseTo(w1, 6);
      expect(h2).toBeCloseTo(h1, 6);

      // Hit testing should use the updated geometry: the old location should no longer hit.
      const hitX2 = x2 + w2 / 2;
      const hitY2 = y2 + h2 / 2;
      expect((app as any).drawingCursorAtPoint(hitX1, hitY1)).toBeNull();
      expect((app as any).drawingCursorAtPoint(hitX2, hitY2)).toBe("move");

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("re-renders drawings when shared-grid row heights change", async () => {
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

      // Disable async IndexedDB hydration so stroke calls land synchronously for this unit test.
      (app as any).drawingImages.getAsync = undefined;

      const drawingCanvas = (app as any).drawingCanvas as HTMLCanvasElement;
      expect(drawingCanvas).toBeTruthy();

      const objects: DrawingObject[] = [
        {
          id: 1,
          kind: { type: "image", imageId: "missing" },
          anchor: {
            type: "oneCell",
            from: { cell: { row: 1, col: 0 }, offset: { xEmu: 0, yEmu: 0 } }, // A2
            size: { cx: pxToEmu(10), cy: pxToEmu(10) },
          },
          zOrder: 0,
        },
      ];

      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();
      calls!.splice(0, calls!.length);

      const renderSpy = vi.spyOn(app as any, "renderDrawings");

      // Initial render.
      (app as any).renderDrawings();
      await new Promise((resolve) => setTimeout(resolve, 0));
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const y1 = Number(firstStroke!.args[1]);
      expect(Number.isFinite(y1)).toBe(true);

      // Resize row 1 (doc row 0 => grid row 1).
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1;
      const prevSize = renderer.getRowHeight(index);
      const nextSize = prevSize + 30;
      renderer.setRowHeight(index, nextSize);

      // Clear previous calls and spy counts.
      calls!.splice(0, calls!.length);
      renderSpy.mockClear();

      (app as any).onSharedGridAxisSizeChange({
        kind: "row",
        index,
        size: nextSize,
        previousSize: prevSize,
        defaultSize: renderer.scroll.rows.defaultSize,
        zoom: renderer.getZoom(),
        source: "resize",
      });
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(renderSpy).toHaveBeenCalled();
      const secondStroke = calls!.find((call) => call.method === "strokeRect");
      expect(secondStroke).toBeTruthy();
      const y2 = Number(secondStroke!.args[1]);
      expect(y2).toBeCloseTo(y1 + (nextSize - prevSize), 6);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("re-renders drawings when shared-grid axis resize is rejected during editing", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const formulaBar = document.createElement("div");
      document.body.appendChild(formulaBar);

      const app = new SpreadsheetApp(root, status, { formulaBar });
      expect(app.getGridMode()).toBe("shared");

      const input = formulaBar.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
      expect(input).not.toBeNull();
      input!.focus();
      input!.value = "=SUM(";
      input!.dispatchEvent(new Event("input", { bubbles: true }));

      // Mimic selecting ranges while the formula bar is still editing.
      root.focus();
      expect(app.isEditing()).toBe(true);

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
      const doc = app.getDocument() as any;
      doc.getSheetDrawings = () => objects;

      const calls = ctxCallsByCanvas.get(drawingCanvas);
      expect(calls).toBeTruthy();
      calls!.splice(0, calls!.length);

      // Baseline render.
      (app as any).renderDrawings();
      const firstStroke = calls!.find((call) => call.method === "strokeRect");
      expect(firstStroke).toBeTruthy();
      const x1 = Number(firstStroke!.args[0]);

      // Mimic interactive drag that changed the renderer but should be rejected because we're editing.
      const sharedGrid = (app as any).sharedGrid;
      const renderer = sharedGrid.renderer;
      const index = 1; // grid col 1 => doc col 0
      const prevSize = renderer.getColWidth(index);
      const nextSize = prevSize + 50;
      renderer.setColWidth(index, nextSize);

      calls!.splice(0, calls!.length);

      (app as any).onSharedGridAxisSizeChange({
        kind: "col",
        index,
        size: nextSize,
        previousSize: prevSize,
        defaultSize: renderer.scroll.cols.defaultSize,
        zoom: renderer.getZoom(),
        source: "resize",
      });

      const secondStroke = calls!.find((call) => call.method === "strokeRect");
      expect(secondStroke).toBeTruthy();
      const x2 = Number(secondStroke!.args[0]);
      expect(x2).toBeCloseTo(x1, 6);

      // Focus should be restored to the formula bar so the user can continue typing.
      expect(document.activeElement).toBe(input);

      app.destroy();
      root.remove();
      formulaBar.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
