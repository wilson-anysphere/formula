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
const PROVIDER_CACHE_MAX_SIZE = 50_000;
const MAX_SHEET_VIEW_KEY_ACCESSES = 200_000;

function getProviderCacheStats(provider: any): { sheetCaches: number; maxSheetSize: number } {
  const sheetCaches: Map<string, any> | undefined = provider?.sheetCaches;
  if (!sheetCaches || typeof sheetCaches.size !== "number") return { sheetCaches: 0, maxSheetSize: 0 };
  let maxSheetSize = 0;
  for (const cache of sheetCaches.values()) {
    const size = typeof cache?.size === "number" ? cache.size : 0;
    if (size > maxSheetSize) maxSheetSize = size;
  }
  return { sheetCaches: sheetCaches.size, maxSheetSize };
}

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

function createNumericKeyCountingProxy<T extends Record<string, number>>(
  obj: T,
  counters: { gets: number; has: number; descriptors: number },
): T {
  const numericKey = /^\d+$/;
  return new Proxy(obj, {
    get(target, prop, receiver) {
      if (typeof prop === "string" && numericKey.test(prop)) counters.gets += 1;
      return Reflect.get(target, prop, receiver);
    },
    has(target, prop) {
      if (typeof prop === "string" && numericKey.test(prop)) counters.has += 1;
      return Reflect.has(target, prop);
    },
    getOwnPropertyDescriptor(target, prop) {
      if (typeof prop === "string" && numericKey.test(prop)) counters.descriptors += 1;
      return Reflect.getOwnPropertyDescriptor(target, prop);
    },
  });
}

function withAllocationGuards(fn: () => void): { elapsedMs: number; mapSetCalls: number; mapGetCalls: number; mapHasCalls: number } {
  const originalArray = globalThis.Array;
  const originalMapSet = Map.prototype.set;
  const originalMapGet = Map.prototype.get;
  const originalMapHas = Map.prototype.has;
  const originalArrayPush = originalArray.prototype.push;
  const originalArrayBuffer = globalThis.ArrayBuffer;
  const typedArrayNames = [
    "Uint8Array",
    "Uint8ClampedArray",
    "Uint16Array",
    "Uint32Array",
    "Int8Array",
    "Int16Array",
    "Int32Array",
    "Float32Array",
    "Float64Array",
  ] as const;
  const originalTypedArrays: Record<string, unknown> = {};
  for (const name of typedArrayNames) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    originalTypedArrays[name] = (globalThis as any)[name];
  }

  const MAX_ARRAY_LENGTH = 200_000;
  const MAX_ARRAY_BUFFER_BYTES = 200_000;
  let mapSetCalls = 0;
  let mapGetCalls = 0;
  let mapHasCalls = 0;
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

  Map.prototype.get = function (...args) {
    mapGetCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapGet as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  Map.prototype.has = function (...args) {
    mapHasCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapHas as any).apply(this, args);
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
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).ArrayBuffer = new Proxy(originalArrayBuffer, {
    construct(target, args, newTarget) {
      const size = typeof args[0] === "number" ? args[0] : NaN;
      if (Number.isFinite(size) && size > MAX_ARRAY_BUFFER_BYTES) {
        throw new Error(`Unexpected large ArrayBuffer allocation: byteLength=${size}`);
      }
      return Reflect.construct(target, args, newTarget);
    },
  });

  for (const name of typedArrayNames) {
    const ctor = originalTypedArrays[name];
    if (!ctor) continue;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any)[name] = new Proxy(ctor as any, {
      construct(target, args, newTarget) {
        const first = args[0];
        if (typeof first === "number" && first > MAX_ARRAY_LENGTH) {
          throw new Error(`Unexpected large ${name} allocation: length=${first}`);
        }
        const byteLength =
          first && typeof first === "object" && typeof (first as any).byteLength === "number" ? (first as any).byteLength : null;
        if (byteLength != null && byteLength > MAX_ARRAY_BUFFER_BYTES) {
          throw new Error(`Unexpected large ${name} backing buffer: byteLength=${byteLength}`);
        }
        return Reflect.construct(target, args, newTarget);
      },
    });
  }

  const start = performance.now();
  try {
    fn();
    return { elapsedMs: performance.now() - start, mapSetCalls, mapGetCalls, mapHasCalls };
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Array = originalArray;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).ArrayBuffer = originalArrayBuffer;
    for (const name of typedArrayNames) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any)[name] = originalTypedArrays[name];
    }
    Map.prototype.set = originalMapSet;
    Map.prototype.get = originalMapGet;
    Map.prototype.has = originalMapHas;
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
      const provider = (app as any).sharedProvider;
      const limits = app.getGridLimits();

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
      const hadOwnGetSheetView = Object.prototype.hasOwnProperty.call(doc, "getSheetView");
      const originalGetSheetView = doc.getSheetView;
      const view = { ...(typeof originalGetSheetView === "function" ? originalGetSheetView.call(doc, sheetId) : {}) } as any;
      const rowViewAccess = { gets: 0, has: 0, descriptors: 0 };
      const colViewAccess = { gets: 0, has: 0, descriptors: 0 };
      if (typeof originalGetSheetView === "function") {
        doc.getSheetView = (id: string) => (id === sheetId ? view : originalGetSheetView.call(doc, id));
      }

      const rowStart = limits.maxRows - OVERRIDE_COUNT - 1;
      const colStart = limits.maxCols - OVERRIDE_COUNT - 1;
      const rowHeights: Record<string, number> = {};
      const colWidths: Record<string, number> = {};
      for (let i = 0; i < OVERRIDE_COUNT; i += 1) {
        rowHeights[String(rowStart + i)] = HIDE_AXIS_SIZE_BASE;
        colWidths[String(colStart + i)] = HIDE_AXIS_SIZE_BASE;
      }

      // "Hide": install sparse overrides for 10k rows/cols.
      view.rowHeights = createNumericKeyCountingProxy(rowHeights, rowViewAccess);
      view.colWidths = createNumericKeyCountingProxy(colWidths, colViewAccess);

      rowViewAccess.gets = 0;
      rowViewAccess.has = 0;
      rowViewAccess.descriptors = 0;
      colViewAccess.gets = 0;
      colViewAccess.has = 0;
      colViewAccess.descriptors = 0;
      const hideRun = withAllocationGuards(() => {
        (app as any).syncSharedGridAxisSizesFromDocument();
      });
      expect(rowViewAccess.gets + rowViewAccess.has + rowViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);
      expect(colViewAccess.gets + colViewAccess.has + colViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);

      const headerRows = (app as any).sharedHeaderRows?.() ?? 1;
      const headerCols = (app as any).sharedHeaderCols?.() ?? 1;
      const zoom = renderer.getZoom();
      expect(renderer.getRowHeight(rowStart + headerRows)).toBeCloseTo(HIDE_AXIS_SIZE_BASE * zoom, 6);
      expect(renderer.getColWidth(colStart + headerCols)).toBeCloseTo(HIDE_AXIS_SIZE_BASE * zoom, 6);

      // Keep bounded by selection size. (Some implementations may store fewer overrides in the future.)
      expect((renderer as any).rowHeightOverridesBase.size).toBeLessThanOrEqual(baselineRowOverrides + OVERRIDE_COUNT);
      expect((renderer as any).colWidthOverridesBase.size).toBeLessThanOrEqual(baselineColOverrides + OVERRIDE_COUNT);

      // Axis override sync should not explode the shared provider cache.
      {
        const stats = getProviderCacheStats(provider);
        expect(stats.sheetCaches).toBeLessThanOrEqual(4);
        expect(stats.maxSheetSize).toBeLessThanOrEqual(PROVIDER_CACHE_MAX_SIZE);
      }

      // "Unhide": clear the overrides and re-sync.
      view.rowHeights = createNumericKeyCountingProxy({}, rowViewAccess);
      view.colWidths = createNumericKeyCountingProxy({}, colViewAccess);

      rowViewAccess.gets = 0;
      rowViewAccess.has = 0;
      rowViewAccess.descriptors = 0;
      colViewAccess.gets = 0;
      colViewAccess.has = 0;
      colViewAccess.descriptors = 0;
      const unhideRun = withAllocationGuards(() => {
        (app as any).syncSharedGridAxisSizesFromDocument();
      });
      expect(rowViewAccess.gets + rowViewAccess.has + rowViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);
      expect(colViewAccess.gets + colViewAccess.has + colViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);

      expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides);
      expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides);

      expect(renderer.getRowHeight(rowStart + headerRows)).toBeCloseTo(renderer.scroll.rows.defaultSize, 6);
      expect(renderer.getColWidth(colStart + headerCols)).toBeCloseTo(renderer.scroll.cols.defaultSize, 6);

      // Two batch sync calls should schedule at most one invalidation per call (not per-index updates).
      expect(requestRenderSpy.mock.calls.length).toBeLessThanOrEqual(2);

      {
        const stats = getProviderCacheStats(provider);
        expect(stats.sheetCaches).toBeLessThanOrEqual(4);
        expect(stats.maxSheetSize).toBeLessThanOrEqual(PROVIDER_CACHE_MAX_SIZE);
      }

      // Guardrails: keep work proportional to the override count, not sheet maxes.
      // Map.set call thresholds are generous and mainly exist to catch accidental
      // `for (row=0..maxRows)` style loops.
      expect(hideRun.mapSetCalls).toBeLessThan(300_000);
      expect(unhideRun.mapSetCalls).toBeLessThan(300_000);
      expect(hideRun.mapGetCalls).toBeLessThan(500_000);
      expect(unhideRun.mapGetCalls).toBeLessThan(500_000);
      expect(hideRun.mapHasCalls).toBeLessThan(500_000);
      expect(unhideRun.mapHasCalls).toBeLessThan(500_000);

      // Perf numbers can fluctuate based on host load (and our test harness spins up a lot of
      // infrastructure). Keep a generous ceiling so we still catch accidental O(maxRows/maxCols)
      // regressions without introducing local flakiness.
      if (!process.env.CI && !process.env.IS_ON_DEV_EC2_MACHINE) {
        expect(hideRun.elapsedMs).toBeLessThan(2_000);
        expect(unhideRun.elapsedMs).toBeLessThan(2_000);
      }

      if (typeof originalGetSheetView === "function") {
        if (hadOwnGetSheetView) doc.getSheetView = originalGetSheetView;
        else delete doc.getSheetView;
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
      const outline = (app as any).getOutlineForSheet(app.getCurrentSheetId()) as any;

      const rowEntries: Map<number, unknown> = outline.rows.entries;
      const colEntries: Map<number, unknown> = outline.cols.entries;
      const hadOwnRowGet = Object.prototype.hasOwnProperty.call(rowEntries, "get");
      const hadOwnColGet = Object.prototype.hasOwnProperty.call(colEntries, "get");
      const hadOwnRowHas = Object.prototype.hasOwnProperty.call(rowEntries, "has");
      const hadOwnColHas = Object.prototype.hasOwnProperty.call(colEntries, "has");
      const baseRowGet = rowEntries.get;
      const baseColGet = colEntries.get;
      const baseRowHas = rowEntries.has;
      const baseColHas = colEntries.has;
      let rowGetCalls = 0;
      let colGetCalls = 0;
      let rowHasCalls = 0;
      let colHasCalls = 0;
      (rowEntries as any).get = function (this: Map<number, unknown>, key: number) {
        rowGetCalls += 1;
        return baseRowGet.call(this, key);
      };
      (colEntries as any).get = function (this: Map<number, unknown>, key: number) {
        colGetCalls += 1;
        return baseColGet.call(this, key);
      };
      (rowEntries as any).has = function (this: Map<number, unknown>, key: number) {
        rowHasCalls += 1;
        return baseRowHas.call(this, key);
      };
      (colEntries as any).has = function (this: Map<number, unknown>, key: number) {
        colHasCalls += 1;
        return baseColHas.call(this, key);
      };

      const restoreOutlineGetOverrides = () => {
        if (hadOwnRowGet) (rowEntries as any).get = baseRowGet;
        else delete (rowEntries as any).get;
        if (hadOwnColGet) (colEntries as any).get = baseColGet;
        else delete (colEntries as any).get;
        if (hadOwnRowHas) (rowEntries as any).has = baseRowHas;
        else delete (rowEntries as any).has;
        if (hadOwnColHas) (colEntries as any).has = baseColHas;
        else delete (colEntries as any).has;
      };

      try {
        expect(app.getGridMode()).toBe("shared");

         const sharedGrid = (app as any).sharedGrid;
         expect(sharedGrid).toBeTruthy();
         const renderer = sharedGrid.renderer;
         const provider = (app as any).sharedProvider;
         const limits = app.getGridLimits();
        const doc = app.getDocument() as any;
        const sheetId = app.getCurrentSheetId();
        const hadOwnGetSheetView = Object.prototype.hasOwnProperty.call(doc, "getSheetView");
        const originalGetSheetView = doc.getSheetView;
        const view = { ...(typeof originalGetSheetView === "function" ? originalGetSheetView.call(doc, sheetId) : {}) } as any;
        const rowViewAccess = { gets: 0, has: 0, descriptors: 0 };
        const colViewAccess = { gets: 0, has: 0, descriptors: 0 };
        if (typeof originalGetSheetView === "function") {
          // Even in the outline-driven hide/unhide path, `syncSharedGridAxisSizesFromDocument` reads
          // `view.rowHeights/colWidths`. Wrap them in proxies so any accidental `for (i=0..maxRows)`
          // scanning loops will trip deterministic counters instead of slipping past allocation-only checks.
          view.rowHeights = createNumericKeyCountingProxy({}, rowViewAccess);
          view.colWidths = createNumericKeyCountingProxy({}, colViewAccess);
          doc.getSheetView = (id: string) => (id === sheetId ? view : originalGetSheetView.call(doc, id));
        }

        const baselineRowOverrides = (renderer as any).rowHeightOverridesBase.size as number;
        const baselineColOverrides = (renderer as any).colWidthOverridesBase.size as number;

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
         const rowStart = limits.maxRows - OVERRIDE_COUNT - 1;
         const colStart = limits.maxCols - OVERRIDE_COUNT - 1;
         const rows: number[] = new Array<number>(OVERRIDE_COUNT);
         const cols: number[] = new Array<number>(OVERRIDE_COUNT);
         for (let i = 0; i < OVERRIDE_COUNT; i += 1) {
           rows[i] = rowStart + i;
           cols[i] = colStart + i;
         }

         rowViewAccess.gets = 0;
         rowViewAccess.has = 0;
         rowViewAccess.descriptors = 0;
         colViewAccess.gets = 0;
         colViewAccess.has = 0;
         colViewAccess.descriptors = 0;
         const hideRun = withAllocationGuards(() => {
           app.hideRows(rows);
           app.hideCols(cols);
         });
         expect(rowViewAccess.gets + rowViewAccess.has + rowViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);
         expect(colViewAccess.gets + colViewAccess.has + colViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);

         expect(rebuildSpy).not.toHaveBeenCalled();
         expect(outline.rows.entries.size).toBeLessThanOrEqual(baselineOutlineRows + OVERRIDE_COUNT);
         expect(outline.cols.entries.size).toBeLessThanOrEqual(baselineOutlineCols + OVERRIDE_COUNT);

        const headerRows = (app as any).sharedHeaderRows?.() ?? 1;
        const headerCols = (app as any).sharedHeaderCols?.() ?? 1;
        const zoom = renderer.getZoom();
        const hiddenAxisSizeBase = 2;
        expect(renderer.getRowHeight(rowStart + headerRows)).toBeCloseTo(hiddenAxisSizeBase * zoom, 6);
        expect(renderer.getColWidth(colStart + headerCols)).toBeCloseTo(hiddenAxisSizeBase * zoom, 6);

        // Keep bounded by selection size. (Future implementations may store fewer overrides.)
        expect((renderer as any).rowHeightOverridesBase.size).toBeLessThanOrEqual(baselineRowOverrides + OVERRIDE_COUNT);
        expect((renderer as any).colWidthOverridesBase.size).toBeLessThanOrEqual(baselineColOverrides + OVERRIDE_COUNT);

        {
          const stats = getProviderCacheStats(provider);
          expect(stats.sheetCaches).toBeLessThanOrEqual(4);
          expect(stats.maxSheetSize).toBeLessThanOrEqual(PROVIDER_CACHE_MAX_SIZE);
        }

         // Ensure the implementation remains sparse: avoid scanning all rows/cols to check hidden state.
         // (Current implementation iterates only `outline.*.entries` plus constant-time checks.)
         expect(rowEntrySpy.mock.calls.length).toBeLessThan(100_000);
         expect(colEntrySpy.mock.calls.length).toBeLessThan(100_000);
         expect(rowGetCalls).toBeLessThan(150_000);
         expect(colGetCalls).toBeLessThan(150_000);
         expect(rowHasCalls).toBeLessThan(150_000);
         expect(colHasCalls).toBeLessThan(150_000);

         rowViewAccess.gets = 0;
         rowViewAccess.has = 0;
         rowViewAccess.descriptors = 0;
         colViewAccess.gets = 0;
         colViewAccess.has = 0;
         colViewAccess.descriptors = 0;
         const unhideRun = withAllocationGuards(() => {
           app.unhideRows(rows);
           app.unhideCols(cols);
         });
         expect(rowViewAccess.gets + rowViewAccess.has + rowViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);
         expect(colViewAccess.gets + colViewAccess.has + colViewAccess.descriptors).toBeLessThan(MAX_SHEET_VIEW_KEY_ACCESSES);

         expect(rebuildSpy).not.toHaveBeenCalled();
         // Unhide should not create additional outline entries; it may keep or delete entries depending on implementation.
         expect(outline.rows.entries.size).toBeLessThanOrEqual(baselineOutlineRows + OVERRIDE_COUNT);
        expect(outline.cols.entries.size).toBeLessThanOrEqual(baselineOutlineCols + OVERRIDE_COUNT);

        expect((renderer as any).rowHeightOverridesBase.size).toBe(baselineRowOverrides);
        expect((renderer as any).colWidthOverridesBase.size).toBe(baselineColOverrides);

        expect(renderer.getRowHeight(rowStart + headerRows)).toBeCloseTo(renderer.scroll.rows.defaultSize, 6);
        expect(renderer.getColWidth(colStart + headerCols)).toBeCloseTo(renderer.scroll.cols.defaultSize, 6);

        // One render invalidation per outline update (hide rows, hide cols, unhide rows, unhide cols),
        // with some tolerance for coalescing / extra selection invalidations in future implementations.
        // (The key regression this catches is per-index invalidation loops.)
        expect(requestRenderSpy.mock.calls.length).toBeLessThanOrEqual(10);

         // Keep work proportional to the number of hidden indices, not sheet maxes.
         expect(hideRun.mapSetCalls).toBeLessThan(600_000);
         expect(unhideRun.mapSetCalls).toBeLessThan(600_000);
         expect(hideRun.mapGetCalls).toBeLessThan(900_000);
         expect(unhideRun.mapGetCalls).toBeLessThan(900_000);
         expect(hideRun.mapHasCalls).toBeLessThan(900_000);
         expect(unhideRun.mapHasCalls).toBeLessThan(900_000);

        {
          const stats = getProviderCacheStats(provider);
          expect(stats.sheetCaches).toBeLessThanOrEqual(4);
          expect(stats.maxSheetSize).toBeLessThanOrEqual(PROVIDER_CACHE_MAX_SIZE);
        }

          // Time-based assertions are intentionally opt-in since wall-clock performance varies wildly
          // across machines / environments (and is especially flaky in shared CI runners).
          // Run with `FORMULA_PERF_ASSERT=1` to enforce a local perf budget.
          if (process.env.FORMULA_PERF_ASSERT === "1") {
            expect(hideRun.elapsedMs).toBeLessThan(1_500);
            expect(unhideRun.elapsedMs).toBeLessThan(1_500);
          }
        } finally {
          if (typeof originalGetSheetView === "function") {
            if (hadOwnGetSheetView) doc.getSheetView = originalGetSheetView;
            else delete doc.getSheetView;
          }
         restoreOutlineGetOverrides();
         app.destroy();
         root.remove();
       }
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
