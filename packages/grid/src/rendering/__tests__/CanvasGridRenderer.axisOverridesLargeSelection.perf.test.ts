/**
 * @vitest-environment jsdom
 *
 * Performance regression coverage for large axis override batches (eg hide/unhide style operations).
 *
 * Goal: ensure applying overrides to a small selection on an Excel-scale sheet stays O(selectionSize),
 * and does not accidentally allocate/work proportional to maxRows/maxCols.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer } from "../CanvasGridRenderer";
import type { CellProvider } from "../../model/CellProvider";

const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLS = 16_384;
const OVERRIDE_COUNT = 10_000;

function buildIndexMap(start: number, count: number, value: number): Map<number, number> {
  const out = new Map<number, number>();
  for (let i = 0; i < count; i += 1) out.set(start + i, value);
  return out;
}

function withAllocationGuards<T>(fn: () => T): {
  result: T;
  elapsedMs: number;
  mapSetCalls: number;
  mapGetCalls: number;
  mapHasCalls: number;
} {
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
    "Float64Array"
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
    }
  });

  Map.prototype.set = function (this: Map<unknown, unknown>, ...args: [unknown, unknown]) {
    mapSetCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapSet as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  Map.prototype.get = function (this: Map<unknown, unknown>, ...args: [unknown]) {
    mapGetCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapGet as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  Map.prototype.has = function (this: Map<unknown, unknown>, ...args: [unknown]) {
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
    }
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
        // If the caller passes an ArrayBuffer directly, validate its size as well.
        const byteLength =
          first && typeof first === "object" && typeof (first as any).byteLength === "number" ? (first as any).byteLength : null;
        if (byteLength != null && byteLength > MAX_ARRAY_BUFFER_BYTES) {
          throw new Error(`Unexpected large ${name} backing buffer: byteLength=${byteLength}`);
        }
        return Reflect.construct(target, args, newTarget);
      }
    });
  }

  const start = performance.now();
  try {
    const result = fn();
    return { result, elapsedMs: performance.now() - start, mapSetCalls, mapGetCalls, mapHasCalls };
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

describe("CanvasGridRenderer axis overrides large-selection perf characteristics", () => {
  beforeEach(() => {
    // Some other tests use fake timers and do not always restore them. This perf regression
    // suite relies on real `performance.now()` measurements, so force real timers here to
    // avoid time-skew flakes when running the full Vitest suite in parallel.
    vi.useRealTimers();
    // Avoid scheduling real animation frames in the unit test environment.
    // (Render cost is not part of this perf regression coverage.)
    vi.stubGlobal("requestAnimationFrame", () => 0);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("applies hide/unhide-style override batches without maxRows/maxCols allocations", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: EXCEL_MAX_ROWS, colCount: EXCEL_MAX_COLS });

    const requestRenderSpy = vi.spyOn(renderer, "requestRender");

    // Apply overrides near the end of the axis to ensure we don't accidentally do work proportional
    // to `maxIndex` (e.g. allocating `new Array(maxIndex)`), not just proportional to the number of
    // overridden indices.
    const rowStart = EXCEL_MAX_ROWS - OVERRIDE_COUNT - 1;
    const colStart = EXCEL_MAX_COLS - OVERRIDE_COUNT - 1;

    // "Hide": set a tiny non-default size for 10k rows/cols.
    // (We use size=1px as a proxy for hidden; the core concern is sparse override batching.)
    const hideRows = buildIndexMap(rowStart, OVERRIDE_COUNT, 1);
    const hideCols = buildIndexMap(colStart, OVERRIDE_COUNT, 1);

    const hideRun = withAllocationGuards(() => {
      renderer.applyAxisSizeOverrides({ rows: hideRows, cols: hideCols }, { resetUnspecified: true });
    });

    // After hide: the override maps should contain only the specified indices (O(10k), not O(maxRows/maxCols)).
    expect((renderer as any).rowHeightOverridesBase.size).toBe(OVERRIDE_COUNT);
    expect((renderer as any).colWidthOverridesBase.size).toBe(OVERRIDE_COUNT);
    expect(renderer.getRowHeight(rowStart)).toBe(1);
    expect(renderer.getColWidth(colStart)).toBe(1);
    expect((renderer.scroll.rows as any).overrides.size).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.cols as any).overrides.size).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.rows as any).overrideIndices.length).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.cols as any).overrideIndices.length).toBe(OVERRIDE_COUNT);

    // "Unhide": clear the overrides by applying an empty override set with `resetUnspecified=true`.
    // This mirrors shared-grid's "re-sync from document state" behavior.
    const unhideRun = withAllocationGuards(() => {
      renderer.applyAxisSizeOverrides({ rows: new Map(), cols: new Map() }, { resetUnspecified: true });
    });

    // Structural assertions:
    // - The persisted override maps must remain O(#overrides), not O(maxRows/maxCols).
    // - Virtual scroll structures must track only overridden indices.
    expect((renderer as any).rowHeightOverridesBase.size).toBe(0);
    expect((renderer as any).colWidthOverridesBase.size).toBe(0);
    expect(renderer.getRowHeight(rowStart)).toBe(renderer.scroll.rows.defaultSize);
    expect(renderer.getColWidth(colStart)).toBe(renderer.scroll.cols.defaultSize);

    const rowsAxis: any = renderer.scroll.rows as any;
    const colsAxis: any = renderer.scroll.cols as any;

    expect(rowsAxis.overrides.size).toBe(0);
    expect(colsAxis.overrides.size).toBe(0);
    expect(rowsAxis.overrideIndices.length).toBe(0);
    expect(colsAxis.overrideIndices.length).toBe(0);
    expect(rowsAxis.diffs.length).toBe(0);
    expect(colsAxis.diffs.length).toBe(0);
    // Fenwick tree arrays are always `diffs.length + 1`.
    expect(rowsAxis.diffBit.length).toBe(1);
    expect(colsAxis.diffBit.length).toBe(1);

    // Each batch apply should trigger a single repaint request (not per-index invalidations).
    // Allow future coalescing (eg if a scheduled frame is already pending).
    expect(requestRenderSpy.mock.calls.length).toBeLessThanOrEqual(2);

    // Guardrails: the override application should not do work proportional to maxRows/maxCols.
    // These are intentionally generous and rely on the allocation guards above for determinism.
    expect(hideRun.mapSetCalls).toBeLessThan(250_000);
    expect(unhideRun.mapSetCalls).toBeLessThan(250_000);
    expect(hideRun.mapGetCalls).toBeLessThan(250_000);
    expect(unhideRun.mapGetCalls).toBeLessThan(250_000);
    expect(hideRun.mapHasCalls).toBeLessThan(250_000);
    expect(unhideRun.mapHasCalls).toBeLessThan(250_000);

    // Time-based assertions are intentionally opt-in since wall-clock performance varies wildly
    // across machines / environments (and is especially flaky in shared CI runners).
    // Run with `FORMULA_PERF_ASSERT=1` to enforce a local perf budget.
    if (process.env.FORMULA_PERF_ASSERT === "1") {
      expect(hideRun.elapsedMs).toBeLessThan(1_000);
      expect(unhideRun.elapsedMs).toBeLessThan(1_000);
    }
  });
});
