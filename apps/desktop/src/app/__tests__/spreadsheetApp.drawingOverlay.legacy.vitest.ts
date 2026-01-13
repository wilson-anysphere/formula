/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DrawingOverlay, pxToEmu } from "../../drawings/overlay";
import { buildHitTestIndex, hitTestDrawings } from "../../drawings/hitTest";
import type { DrawingObject, ImageStore } from "../../drawings/types";
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

describe("SpreadsheetApp drawing overlay (legacy grid)", () => {
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

    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 2 });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("mounts between the base + content canvases and resizes with DPR", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const resizeSpy = vi.spyOn(DrawingOverlay.prototype, "resize");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("legacy");

      const drawingCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      expect(drawingCanvas).not.toBeNull();

      const gridCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--base");
      const contentCanvas = root.querySelector<HTMLCanvasElement>("canvas.grid-canvas--content");
      expect(gridCanvas).not.toBeNull();
      expect(contentCanvas).not.toBeNull();

      const children = Array.from(root.children);
      const gridIdx = children.indexOf(gridCanvas!);
      const drawingIdx = children.indexOf(drawingCanvas!);
      const contentIdx = children.indexOf(contentCanvas!);
      expect(gridIdx).toBeGreaterThanOrEqual(0);
      expect(drawingIdx).toBeGreaterThan(gridIdx);
      expect(drawingIdx).toBeLessThan(contentIdx);

      expect(resizeSpy).toHaveBeenCalledWith(
        expect.objectContaining({
          width: 800 - 48,
          height: 600 - 24,
          dpr: 2,
        }),
      );
      expect(drawingCanvas!.width).toBe((800 - 48) * 2);
      expect(drawingCanvas!.height).toBe((600 - 24) * 2);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("passes frozen pane counts from the document even if derived frozen counts are stale", () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const renderSpy = vi.spyOn(DrawingOverlay.prototype, "render");

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      // Simulate stale cached counts (charts historically had a bug here).
      (app as any).frozenRows = 0;
      (app as any).frozenCols = 0;
      (app as any).scrollX = 50;
      (app as any).scrollY = 100;

      renderSpy.mockClear();
      app.refresh("scroll");

      expect(renderSpy).toHaveBeenCalled();
      const lastCall = renderSpy.mock.calls.at(-1);
      const viewport = lastCall?.[1] as any;

      expect(viewport).toEqual(
        expect.objectContaining({
          frozenRows: 1,
          frozenCols: 1,
        }),
      );

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("computes consistent render vs interaction viewports for drawings (legacy grid)", async () => {
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
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });
      app.setScroll(50, 10);

      const renderViewport = app.getDrawingRenderViewport();
      const interactionViewport = app.getDrawingInteractionViewport();

      expect(renderViewport.headerOffsetX).toBe(0);
      expect(renderViewport.headerOffsetY).toBe(0);
      expect(interactionViewport.headerOffsetX).toBeGreaterThan(0);
      expect(interactionViewport.headerOffsetY).toBeGreaterThan(0);

      // Frozen boundaries should map between viewport spaces by subtracting header offsets.
      expect(interactionViewport.frozenWidthPx! - interactionViewport.headerOffsetX!).toBe(renderViewport.frozenWidthPx);
      expect(interactionViewport.frozenHeightPx! - interactionViewport.headerOffsetY!).toBe(renderViewport.frozenHeightPx);

      // Verify hit testing aligns with where the object is rendered in drawingCanvas space.
      const geom = (app as any).drawingOverlay.geom;
      const images: ImageStore = { get: () => undefined, set: () => {} };

      const object: DrawingObject = {
        id: 1,
        kind: { type: "shape" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 2, col: 2 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: pxToEmu(50), cy: pxToEmu(20) },
        },
        zOrder: 0,
      };

      const calls: Array<{ method: string; args: unknown[] }> = [];
      const ctx: any = {
        clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
        save: () => calls.push({ method: "save", args: [] }),
        restore: () => calls.push({ method: "restore", args: [] }),
        beginPath: () => calls.push({ method: "beginPath", args: [] }),
        rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
        clip: () => calls.push({ method: "clip", args: [] }),
        setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
        strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
        fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
      };

      const canvas: any = {
        width: 0,
        height: 0,
        style: {},
        getContext: () => ctx,
      };

      const overlay = new DrawingOverlay(canvas as HTMLCanvasElement, images, geom);
      await overlay.render([object], renderViewport);

      const stroke = calls.find((c) => c.method === "strokeRect");
      expect(stroke).toBeTruthy();
      const [renderX, renderY, w, h] = stroke!.args as number[];

      const index = buildHitTestIndex([object], geom, { bucketSizePx: 64 });
      const hit = hitTestDrawings(
        index,
        interactionViewport,
        renderX + interactionViewport.headerOffsetX! + 1,
        renderY + interactionViewport.headerOffsetY! + 1,
      );
      expect(hit?.object.id).toBe(1);
      expect(hit?.bounds).toEqual({
        x: renderX + interactionViewport.headerOffsetX!,
        y: renderY + interactionViewport.headerOffsetY!,
        width: w,
        height: h,
      });

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
