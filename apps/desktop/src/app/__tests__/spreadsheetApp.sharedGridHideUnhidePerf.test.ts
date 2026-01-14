/**
 * @vitest-environment jsdom
 *
 * Performance regression coverage for shared-grid hide/unhide-style operations.
 *
 * Goal: avoid O(maxRows/maxCols) allocations when applying sparse row/col overrides to
 * large (Excel-scale) sheets.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

const OVERRIDE_COUNT = 10_000;
const HIDE_AXIS_SIZE_BASE = 1;

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

function withAllocationGuards(fn: () => void): { elapsedMs: number; mapSetCalls: number } {
  const originalArray = globalThis.Array;
  const originalMapSet = Map.prototype.set;
  const originalArrayPush = originalArray.prototype.push;

  const MAX_ARRAY_LENGTH = 200_000;
  let mapSetCalls = 0;
  let pushedElements = 0;

  const GuardedArray = new Proxy(originalArray, {
    get(target, prop, receiver) {
      // Catch `Array.from({ length: N })` patterns which allocate based on the `length` property.
      if (prop === "from") {
        return function guardedFrom(arrayLike: unknown, mapFn?: unknown, thisArg?: unknown) {
          const length =
            arrayLike && (typeof arrayLike === "object" || typeof arrayLike === "function")
              ? // eslint-disable-next-line @typescript-eslint/no-explicit-any
                Number((arrayLike as any).length)
              : NaN;
          if (Number.isFinite(length) && length > MAX_ARRAY_LENGTH) {
            throw new Error(`Unexpected large Array allocation via Array.from: length=${length}`);
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          return (originalArray.from as any)(arrayLike as any, mapFn as any, thisArg as any);
        };
      }
      return Reflect.get(target, prop, receiver);
    },
    construct(target, args) {
      if (args.length === 1 && typeof args[0] === "number" && args[0] > MAX_ARRAY_LENGTH) {
        throw new Error(`Unexpected large Array allocation: length=${args[0]}`);
      }
      return Reflect.construct(target, args);
    },
    apply(target, thisArg, args) {
      if (args.length === 1 && typeof args[0] === "number" && args[0] > MAX_ARRAY_LENGTH) {
        throw new Error(`Unexpected large Array allocation: length=${args[0]}`);
      }
      return Reflect.apply(target, thisArg, args);
    },
  });

  Map.prototype.set = function (...args) {
    mapSetCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapSet as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  // Guard against incremental construction of huge arrays via `arr.push(...)` in a loop.
  originalArray.prototype.push = function (...args) {
    pushedElements += args.length;
    if (pushedElements > MAX_ARRAY_LENGTH) {
      throw new Error(`Unexpected large Array growth via push: pushedElements=${pushedElements}`);
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalArrayPush as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).Array = GuardedArray;

  const start = performance.now();
  try {
    fn();
    return { elapsedMs: performance.now() - start, mapSetCalls };
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Array = originalArray;
    Map.prototype.set = originalMapSet;
    originalArray.prototype.push = originalArrayPush;
  }
}

describe("SpreadsheetApp shared-grid hide/unhide perf", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  afterEach(() => {
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      writable: true,
      value: originalGetContext,
    });
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
    // Provide a stable in-memory implementation for unit tests.
    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // The perf coverage here focuses on the axis override batch plumbing, not actual painting.
    // Keep requestAnimationFrame cheap and deterministic by making it a no-op.
    vi.stubGlobal("requestAnimationFrame", () => 0);
    vi.stubGlobal("cancelAnimationFrame", () => {});

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      writable: true,
      value: () => createMockCanvasContext(),
    });

    vi.stubGlobal(
      "ResizeObserver",
      class {
        observe() {}
        disconnect() {}
      },
    );
  });

  it("applies and clears 10k row/col overrides without O(maxRows/maxCols) work", () => {
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

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      const baselineRowOverrides = (renderer as any).rowHeightOverridesBase.size as number;
      const baselineColOverrides = (renderer as any).colWidthOverridesBase.size as number;
      // Baseline should only include fixed header overrides (eg row-header column width), never
      // anything proportional to sheet maxes.
      expect(baselineRowOverrides).toBeLessThanOrEqual(4);
      expect(baselineColOverrides).toBeLessThanOrEqual(4);

      const requestRenderSpy = vi.spyOn(renderer, "requestRender");
      requestRenderSpy.mockClear();

      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument() as any;
      const view = doc.getSheetView(sheetId) ?? {};

      const rowHeights: Record<string, number> = {};
      const colWidths: Record<string, number> = {};
      for (let i = 0; i < OVERRIDE_COUNT; i += 1) {
        rowHeights[String(i)] = HIDE_AXIS_SIZE_BASE;
        colWidths[String(i)] = HIDE_AXIS_SIZE_BASE;
      }

      // "Hide": install sparse overrides for 10k rows/cols.
      (view as any).rowHeights = rowHeights;
      (view as any).colWidths = colWidths;
      doc.model.setSheetView(sheetId, view);

      const hideRun = withAllocationGuards(() => {
        (app as any).syncSharedGridAxisSizesFromDocument();
      });

      expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides + OVERRIDE_COUNT);
      expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides + OVERRIDE_COUNT);

      // "Unhide": clear the overrides and re-sync.
      (view as any).rowHeights = {};
      (view as any).colWidths = {};
      doc.model.setSheetView(sheetId, view);

      const unhideRun = withAllocationGuards(() => {
        (app as any).syncSharedGridAxisSizesFromDocument();
      });

      expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides);
      expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides);

      // Two batch sync calls => two invalidations (one per operation), not per-index updates.
      expect(requestRenderSpy).toHaveBeenCalledTimes(2);

      // Guardrails: keep work proportional to the override count, not sheet maxes.
      // Map.set call thresholds are generous and mainly exist to catch accidental
      // `for (row=0..maxRows)` style loops.
      expect(hideRun.mapSetCalls).toBeLessThan(300_000);
      expect(unhideRun.mapSetCalls).toBeLessThan(300_000);

      if (!process.env.CI) {
        expect(hideRun.elapsedMs).toBeLessThan(1_000);
        expect(unhideRun.elapsedMs).toBeLessThan(1_000);
      }

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("hides/unhides 10k rows+cols via shared-grid outline without O(maxRows/maxCols) work", () => {
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

      const sharedGrid = (app as any).sharedGrid;
      expect(sharedGrid).toBeTruthy();
      const renderer = sharedGrid.renderer;

      const baselineRowOverrides = (renderer as any).rowHeightOverridesBase.size as number;
      const baselineColOverrides = (renderer as any).colWidthOverridesBase.size as number;

      const outline = (app as any).getOutlineForSheet(app.getCurrentSheetId()) as any;
      const baselineOutlineRows = outline.rows.entries.size as number;
      const baselineOutlineCols = outline.cols.entries.size as number;

      const rebuildSpy = vi.spyOn(app as any, "rebuildAxisVisibilityCache");
      const rowEntrySpy = vi.spyOn(outline.rows, "entry");
      const colEntrySpy = vi.spyOn(outline.cols, "entry");
      rebuildSpy.mockClear();
      rowEntrySpy.mockClear();
      colEntrySpy.mockClear();

      const requestRenderSpy = vi.spyOn(renderer, "requestRender");
      requestRenderSpy.mockClear();

      // Hide a block far away from the active cell so `ensureActiveCellVisible` / `scrollCellIntoView`
      // should not trigger additional scroll/selection work.
      const rowStart = 20_000;
      const colStart = 2_000;
      const rows: number[] = new Array<number>(OVERRIDE_COUNT);
      const cols: number[] = new Array<number>(OVERRIDE_COUNT);
      for (let i = 0; i < OVERRIDE_COUNT; i += 1) {
        rows[i] = rowStart + i;
        cols[i] = colStart + i;
      }

      const hideRun = withAllocationGuards(() => {
        app.hideRows(rows);
        app.hideCols(cols);
      });

      expect(rebuildSpy).not.toHaveBeenCalled();
      expect(outline.rows.entries.size).toBe(baselineOutlineRows + OVERRIDE_COUNT);
      expect(outline.cols.entries.size).toBe(baselineOutlineCols + OVERRIDE_COUNT);

      expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides + OVERRIDE_COUNT);
      expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides + OVERRIDE_COUNT);

      // Ensure the implementation remains sparse: avoid scanning all rows/cols to check hidden state.
      // (Current implementation iterates only `outline.*.entries` plus constant-time checks.)
      expect(rowEntrySpy.mock.calls.length).toBeLessThan(100_000);
      expect(colEntrySpy.mock.calls.length).toBeLessThan(100_000);

      const unhideRun = withAllocationGuards(() => {
        app.unhideRows(rows);
        app.unhideCols(cols);
      });

      expect(rebuildSpy).not.toHaveBeenCalled();
      // Unhide should not create additional outline entries; it just toggles `hidden.user`.
      expect(outline.rows.entries.size).toBe(baselineOutlineRows + OVERRIDE_COUNT);
      expect(outline.cols.entries.size).toBe(baselineOutlineCols + OVERRIDE_COUNT);

      expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides);
      expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides);

      // One render invalidation per outline update (hide rows, hide cols, unhide rows, unhide cols).
      expect(requestRenderSpy).toHaveBeenCalledTimes(4);

      // Keep work proportional to the number of hidden indices, not sheet maxes.
      expect(hideRun.mapSetCalls).toBeLessThan(600_000);
      expect(unhideRun.mapSetCalls).toBeLessThan(600_000);

      if (!process.env.CI) {
        expect(hideRun.elapsedMs).toBeLessThan(1_500);
        expect(unhideRun.elapsedMs).toBeLessThan(1_500);
      }

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
