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

function createMockCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas,
      // Spies we assert on.
      rect: vi.fn(),
      clip: vi.fn(),
      strokeRect: vi.fn(),
      clearRect: vi.fn(),
      // CanvasGridRenderer expects some text measurement APIs to exist.
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

function createShape(opts: { id: number; row: number; col: number; widthPx: number; heightPx: number; zOrder: number }): DrawingObject {
  return {
    id: opts.id,
    kind: { type: "shape" },
    anchor: {
      type: "oneCell",
      from: { cell: { row: opts.row, col: opts.col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(opts.widthPx), cy: pxToEmu(opts.heightPx) },
    },
    zOrder: opts.zOrder,
  };
}

describe("SpreadsheetApp drawings + frozen panes (shared grid)", () => {
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

    // Reuse a stable 2d context per canvas so the app and DrawingOverlay share the same spies.
    const ctxByCanvas = new WeakMap<HTMLCanvasElement, CanvasRenderingContext2D>();
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value(this: HTMLCanvasElement, type: string) {
        if (type !== "2d") return null;
        const existing = ctxByCanvas.get(this);
        if (existing) return existing;
        const ctx = createMockCanvasContext(this);
        ctxByCanvas.set(this, ctx);
        return ctx;
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("clips drawings to frozen-pane quadrants (Excel object pane behavior)", async () => {
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

      // Freeze 1 row + 1 col (sheet space) via the DocumentController.
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const sharedGrid = (app as any).sharedGrid;

      // Set scroll offsets to ensure per-pane scroll behavior is exercised.
      sharedGrid.scrollTo(10, 10);
      const sharedViewport = sharedGrid.renderer.scroll.getViewportState();
      const renderViewport = app.getDrawingRenderViewport(sharedViewport);

      const scrollX = renderViewport.scrollX;
      const scrollY = renderViewport.scrollY;
      const frozenContentWidth = renderViewport.frozenWidthPx ?? 0;
      const frozenContentHeight = renderViewport.frozenHeightPx ?? 0;
      const scrollableWidth = Math.max(0, renderViewport.width - frozenContentWidth);
      const scrollableHeight = Math.max(0, renderViewport.height - frozenContentHeight);

      // Create one drawing per quadrant.
      const objects: DrawingObject[] = [
        // Top-left: intentionally spans across the freeze lines (should be clipped).
        createShape({
          id: 1,
          row: 0,
          col: 0,
          widthPx: frozenContentWidth + 40,
          heightPx: frozenContentHeight + 40,
          zOrder: 0,
        }),
        createShape({ id: 2, row: 0, col: 1, widthPx: 20, heightPx: 10, zOrder: 1 }), // top-right
        createShape({ id: 3, row: 1, col: 0, widthPx: 20, heightPx: 10, zOrder: 2 }), // bottom-left
        createShape({ id: 4, row: 1, col: 1, widthPx: 20, heightPx: 10, zOrder: 3 }), // bottom-right
      ];
      (doc as any).getSheetDrawings = (sheetId: string) => (sheetId === app.getCurrentSheetId() ? objects : []);

      // Clear any calls from initial layout passes.
      const drawingsCanvas = root.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]')!;
      const drawingsCtx = drawingsCanvas.getContext("2d") as CanvasRenderingContext2D & {
        rect: ReturnType<typeof vi.fn>;
        clip: ReturnType<typeof vi.fn>;
        strokeRect: ReturnType<typeof vi.fn>;
      };
      drawingsCtx.rect.mockClear();
      drawingsCtx.clip.mockClear();
      drawingsCtx.strokeRect.mockClear();

      // Force a render pass with the current shared-grid viewport.
      (app as any).renderDrawings(sharedViewport);

      // Four panes should be clipped (one per quadrant).
      expect(drawingsCtx.rect).toHaveBeenCalledTimes(4);
      expect(drawingsCtx.clip).toHaveBeenCalledTimes(4);

      const expectedClipRects = [
        { x: 0, y: 0, width: frozenContentWidth, height: frozenContentHeight },
        { x: frozenContentWidth, y: 0, width: scrollableWidth, height: frozenContentHeight },
        { x: 0, y: frozenContentHeight, width: frozenContentWidth, height: scrollableHeight },
        {
          x: frozenContentWidth,
          y: frozenContentHeight,
          width: scrollableWidth,
          height: scrollableHeight,
        },
      ];

      const rectCalls = drawingsCtx.rect.mock.calls.map((args) => ({
        x: args[0] as number,
        y: args[1] as number,
        width: args[2] as number,
        height: args[3] as number,
      }));
      expect(rectCalls).toEqual(expectedClipRects);

      // Verify per-pane scroll offsets by inspecting the placeholder strokeRects.
      expect(drawingsCtx.strokeRect).toHaveBeenCalledTimes(4);
      const strokeCalls = drawingsCtx.strokeRect.mock.calls.map((args) => ({
        x: args[0] as number,
        y: args[1] as number,
      }));
      expect(strokeCalls[0]).toEqual({ x: 0, y: 0 }); // top-left: no scroll
      expect(strokeCalls[1]).toEqual({ x: frozenContentWidth - scrollX, y: 0 }); // top-right: scrollX only
      expect(strokeCalls[2]).toEqual({ x: 0, y: frozenContentHeight - scrollY }); // bottom-left: scrollY only
      expect(strokeCalls[3]).toEqual({
        x: frozenContentWidth - scrollX,
        y: frozenContentHeight - scrollY,
      }); // bottom-right: scrollX+scrollY

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
