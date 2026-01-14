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
    const priorCanvasCharts = process.env.CANVAS_CHARTS;
    process.env.DESKTOP_GRID_MODE = "shared";
    // Disable canvas charts so the demo ChartStore chart does not add extra draw objects
    // (and additional nested clip() calls) to this drawing-overlay unit test.
    process.env.CANVAS_CHARTS = "0";
    try {
      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      expect(app.getGridMode()).toBe("shared");

      // SpreadsheetApp seeds a demo ChartStore chart on startup in non-collab mode. This test
      // asserts frozen-pane clipping for workbook drawings; clear the seeded chart so it doesn't
      // add extra clip rects/stroke calls to the drawings overlay.
      for (const chart of app.listCharts()) {
        (app as any).chartStore.deleteChart(chart.id);
      }

      // Freeze 1 row + 1 col (sheet space) via the DocumentController.
      const doc = app.getDocument();
      doc.setFrozen(app.getCurrentSheetId(), 1, 1, { label: "Freeze" });

      const sharedGrid = (app as any).sharedGrid;

      // Set scroll offsets to ensure per-pane scroll behavior is exercised.
      sharedGrid.scrollTo(10, 10);
      const sharedViewport = sharedGrid.renderer.scroll.getViewportState();
      const renderViewport = app.getDrawingRenderViewport(sharedViewport);

      const headerOffsetX = renderViewport.headerOffsetX ?? 0;
      const headerOffsetY = renderViewport.headerOffsetY ?? 0;

      const scrollX = renderViewport.scrollX;
      const scrollY = renderViewport.scrollY;
      const frozenBoundaryX = renderViewport.frozenWidthPx ?? headerOffsetX;
      const frozenBoundaryY = renderViewport.frozenHeightPx ?? headerOffsetY;
      const frozenContentWidth = Math.max(0, frozenBoundaryX - headerOffsetX);
      const frozenContentHeight = Math.max(0, frozenBoundaryY - headerOffsetY);
      const cellAreaWidth = Math.max(0, renderViewport.width - headerOffsetX);
      const cellAreaHeight = Math.max(0, renderViewport.height - headerOffsetY);
      const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
      const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

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
      // Note: in canvas-charts mode, the chart renderer can also use `rect`/`clip` internally.
      // Assert on the *pane* clip rects (first 4 calls) rather than the total call count.
      expect(drawingsCtx.rect.mock.calls.length).toBeGreaterThanOrEqual(4);
      expect(drawingsCtx.clip.mock.calls.length).toBeGreaterThanOrEqual(4);
      const expectedClipRects = [
        { x: headerOffsetX, y: headerOffsetY, width: frozenContentWidth, height: frozenContentHeight },
        {
          x: headerOffsetX + frozenContentWidth,
          y: headerOffsetY,
          width: scrollableWidth,
          height: frozenContentHeight,
        },
        {
          x: headerOffsetX,
          y: headerOffsetY + frozenContentHeight,
          width: frozenContentWidth,
          height: scrollableHeight,
        },
        {
          x: headerOffsetX + frozenContentWidth,
          y: headerOffsetY + frozenContentHeight,
          width: scrollableWidth,
          height: scrollableHeight,
        },
      ];

      const rectCalls = drawingsCtx.rect.mock.calls.slice(0, 4).map((args) => ({
        x: args[0] as number,
        y: args[1] as number,
        width: args[2] as number,
        height: args[3] as number,
      }));
      const paneRectCalls = rectCalls.filter((call) =>
        expectedClipRects.some(
          (expected) =>
            call.x === expected.x && call.y === expected.y && call.width === expected.width && call.height === expected.height,
        ),
      );
      const seen = new Set<string>();
      const dedupedPaneRectCalls = paneRectCalls.filter((call) => {
        const key = `${call.x},${call.y},${call.width},${call.height}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      expect(dedupedPaneRectCalls).toEqual(expectedClipRects);

      // Verify per-pane scroll offsets by inspecting the placeholder strokeRects.
      expect(drawingsCtx.strokeRect.mock.calls.length).toBeGreaterThanOrEqual(4);
      const strokeCalls = drawingsCtx.strokeRect.mock.calls.slice(0, 4).map((args) => ({
        x: args[0] as number,
        y: args[1] as number,
      }));
      const expectedStrokeCalls = [
        { x: headerOffsetX, y: headerOffsetY }, // top-left: no scroll
        { x: headerOffsetX + frozenContentWidth - scrollX, y: headerOffsetY }, // top-right: scrollX only
        { x: headerOffsetX, y: headerOffsetY + frozenContentHeight - scrollY }, // bottom-left: scrollY only
        {
          x: headerOffsetX + frozenContentWidth - scrollX,
          y: headerOffsetY + frozenContentHeight - scrollY,
        }, // bottom-right: scrollX+scrollY
      ];
      for (const expected of expectedStrokeCalls) {
        expect(strokeCalls).toContainEqual(expected);
      }

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
      if (priorCanvasCharts === undefined) delete process.env.CANVAS_CHARTS;
      else process.env.CANVAS_CHARTS = priorCanvasCharts;
    }
  });
});
