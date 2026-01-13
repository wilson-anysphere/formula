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

function buildIndexMap(size: number, value: number): Map<number, number> {
  const out = new Map<number, number>();
  for (let i = 0; i < size; i += 1) out.set(i, value);
  return out;
}

function withAllocationGuards<T>(fn: () => T): { result: T; elapsedMs: number; mapSetCalls: number } {
  const originalArray = globalThis.Array;
  const originalMapSet = Map.prototype.set;

  const MAX_ARRAY_LENGTH = 200_000;
  let mapSetCalls = 0;

  const GuardedArray = new Proxy(originalArray, {
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

  Map.prototype.set = function (...args) {
    mapSetCalls += 1;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (originalMapSet as any).apply(this, args);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).Array = GuardedArray;

  const start = performance.now();
  try {
    const result = fn();
    return { result, elapsedMs: performance.now() - start, mapSetCalls };
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).Array = originalArray;
    Map.prototype.set = originalMapSet;
  }
}

describe("CanvasGridRenderer axis overrides large-selection perf characteristics", () => {
  beforeEach(() => {
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

    // "Hide": set a tiny non-default size for 10k rows/cols.
    // (We use size=1px as a proxy for hidden; the core concern is sparse override batching.)
    const hideRows = buildIndexMap(OVERRIDE_COUNT, 1);
    const hideCols = buildIndexMap(OVERRIDE_COUNT, 1);

    const hideRun = withAllocationGuards(() => {
      renderer.applyAxisSizeOverrides({ rows: hideRows, cols: hideCols }, { resetUnspecified: false });
    });

    // After hide: the override maps should contain only the specified indices (O(10k), not O(maxRows/maxCols)).
    expect((renderer as any).rowHeightOverridesBase.size).toBe(OVERRIDE_COUNT);
    expect((renderer as any).colWidthOverridesBase.size).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.rows as any).overrides.size).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.cols as any).overrides.size).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.rows as any).overrideIndices.length).toBe(OVERRIDE_COUNT);
    expect((renderer.scroll.cols as any).overrideIndices.length).toBe(OVERRIDE_COUNT);

    // "Unhide": clear the overrides by applying default-sized entries for the same indices.
    const defaultRow = renderer.scroll.rows.defaultSize;
    const defaultCol = renderer.scroll.cols.defaultSize;
    const unhideRows = buildIndexMap(OVERRIDE_COUNT, defaultRow);
    const unhideCols = buildIndexMap(OVERRIDE_COUNT, defaultCol);

    const unhideRun = withAllocationGuards(() => {
      renderer.applyAxisSizeOverrides({ rows: unhideRows, cols: unhideCols }, { resetUnspecified: false });
    });

    // Structural assertions:
    // - The persisted override maps must remain O(#overrides), not O(maxRows/maxCols).
    // - Virtual scroll structures must track only overridden indices.
    expect((renderer as any).rowHeightOverridesBase.size).toBe(0);
    expect((renderer as any).colWidthOverridesBase.size).toBe(0);

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
    expect(requestRenderSpy).toHaveBeenCalledTimes(2);

    // Guardrails: the override application should not do work proportional to maxRows/maxCols.
    // These are intentionally generous and rely on the allocation guards above for determinism.
    expect(hideRun.mapSetCalls).toBeLessThan(250_000);
    expect(unhideRun.mapSetCalls).toBeLessThan(250_000);

    if (!process.env.CI) {
      expect(hideRun.elapsedMs).toBeLessThan(1_000);
      expect(unhideRun.elapsedMs).toBeLessThan(1_000);
    }
  });
});
