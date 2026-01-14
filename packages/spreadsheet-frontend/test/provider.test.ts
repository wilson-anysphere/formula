import { describe, expect, it } from "vitest";
import type { CellChange, CellData as EngineCellData, CellScalar, EngineClient } from "@formula/engine";
import type { CellProviderUpdate } from "@formula/grid";
import { EngineCellCache, EngineGridProvider, fromA1, toA1 } from "../src/index.js";

class FakeEngine {
  calls: Array<{ range: string; sheet?: string }> = [];

  constructor(
    private readonly values: Map<string, CellScalar>,
    private readonly recalcChanges: CellChange[] = []
  ) {}

  async getRange(range: string, sheet?: string): Promise<EngineCellData[][]> {
    this.calls.push({ range, sheet });

    const [start, end = start] = range.split(":");
    const startPos = fromA1(start);
    const endPos = fromA1(end);

    const startRow0 = Math.min(startPos.row0, endPos.row0);
    const endRow0 = Math.max(startPos.row0, endPos.row0);
    const startCol0 = Math.min(startPos.col0, endPos.col0);
    const endCol0 = Math.max(startPos.col0, endPos.col0);

    const sheetName = sheet ?? "Sheet1";
    const rows: EngineCellData[][] = [];
    for (let r = startRow0; r <= endRow0; r++) {
      const row: EngineCellData[] = [];
      for (let c = startCol0; c <= endCol0; c++) {
        const address = toA1(r, c);
        const key = `${sheetName}!${address}`;
        const value = this.values.get(key) ?? null;
        row.push({ sheet: sheetName, address, input: value, value });
      }
      rows.push(row);
    }
    return rows;
  }

  async recalculate(_sheet?: string): Promise<CellChange[]> {
    return this.recalcChanges;
  }
}

async function flushMicrotasks(times = 3): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("EngineGridProvider", () => {
  it("prefetch populates cache and notifies subscribers", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", true);
    values.set("Sheet1!B1", 42);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 2 });
    await flushMicrotasks();

    expect(cache.getValue(0, 0)).toBe(true);
    expect(cache.getValue(0, 1)).toBe(42);
    expect(provider.getCell(0, 0)?.value).toBe(true);
    expect(provider.getCell(0, 1)?.value).toBe(42);

    expect(updates).toEqual([{ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 } }]);
  });

  it("passes the configured sheet name through to engine range calls", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet2!A1", 99);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10, sheet: "Sheet2" });

    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

    expect(engine.calls).toEqual([{ range: "A1", sheet: "Sheet2" }]);
    expect(provider.getCell(0, 0)?.value).toBe(99);
  });

  it("does not re-fetch or emit updates for cached ranges", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", 1);
    values.set("Sheet1!B1", 2);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 2 });
    await flushMicrotasks();
    expect(engine.calls).toHaveLength(1);
    expect(updates).toHaveLength(1);

    updates.length = 0;
    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 2 });
    await flushMicrotasks();
    expect(engine.calls).toHaveLength(1);
    expect(updates).toEqual([]);
  });

  it("batches prefetch calls in the same microtask", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", 1);
    values.set("Sheet1!B1", 2);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    const p1 = provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    const p2 = provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 1, endCol: 2 });
    await Promise.all([p1, p2]);
    await flushMicrotasks();

    expect(engine.calls.map((c) => c.range)).toEqual(["A1:B1"]);
    expect(provider.getCell(0, 0)?.value).toBe(1);
    expect(provider.getCell(0, 1)?.value).toBe(2);
    expect(updates).toEqual([{ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 } }]);
  });

  it("does not merge disjoint prefetch ranges into a huge bounding box", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", 1);
    values.set("Sheet1!A101", 2);

    const engine = new FakeEngine(values) as any;
    const cache = new EngineCellCache(engine);
    const provider = new EngineGridProvider({ cache, rowCount: 1_000, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    const p1 = provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    const p2 = provider.prefetchAsync({ startRow: 100, endRow: 101, startCol: 0, endCol: 1 });
    await Promise.all([p1, p2]);
    await flushMicrotasks();

    expect(engine.calls.map((c: { range: string }) => c.range).sort()).toEqual(["A1", "A101"]);

    expect(provider.getCell(0, 0)?.value).toBe(1);
    expect(provider.getCell(100, 0)?.value).toBe(2);

    expect(updates).toHaveLength(2);
    const ranges = updates.map((u) => (u.type === "cells" ? u.range : null)).filter(Boolean);
    expect(ranges).toContainEqual({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    expect(ranges).toContainEqual({ startRow: 100, endRow: 101, startCol: 0, endCol: 1 });
  });

  it("supports header row/col offset mode", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", 1);
    values.set("Sheet1!B1", 3);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 100, colCount: 100, headers: true });

    expect(provider.getCell(0, 0)?.value).toBeNull();
    expect(provider.getCell(0, 1)?.value).toBe("A");
    expect(provider.getCell(0, 2)?.value).toBe("B");
    expect(provider.getCell(1, 0)?.value).toBe(1);

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    await provider.prefetchAsync({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });
    await flushMicrotasks();

    expect(engine.calls.map((c) => c.range)).toEqual(["A1:B1"]);
    expect(provider.getCell(1, 1)?.value).toBe(1);
    expect(provider.getCell(1, 2)?.value).toBe(3);

    expect(updates).toEqual([{ type: "cells", range: { startRow: 1, endRow: 2, startCol: 1, endCol: 3 } }]);
  });

  it("can switch sheets by clearing the cache and invalidating the viewport", async () => {
    const values = new Map<string, CellScalar>();
    values.set("Sheet1!A1", 1);
    values.set("Sheet2!A1", 2);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10, sheet: "Sheet1" });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    await flushMicrotasks();
    expect(provider.getCell(0, 0)?.value).toBe(1);

    updates.length = 0;
    provider.setSheet("Sheet2");
    expect(provider.getCell(0, 0)?.value).toBeNull();
    expect(updates).toEqual([{ type: "invalidateAll" }]);

    updates.length = 0;
    await provider.prefetchAsync({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    await flushMicrotasks();
    expect(provider.getCell(0, 0)?.value).toBe(2);
    expect(engine.calls.at(-1)).toEqual({ range: "A1", sheet: "Sheet2" });
    expect(updates).toEqual([{ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 } }]);
  });

  it("coalesces adjacent invalidations when applying recalc changes", async () => {
    const engine = new FakeEngine(new Map());
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    const changes: CellChange[] = [
      { sheet: "Sheet1", address: "A1", value: 1 },
      { sheet: "Sheet1", address: "B1", value: 2 }
    ];
    provider.applyRecalcChanges(changes);
    await flushMicrotasks();

    expect(updates).toEqual([{ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 } }]);
  });

  it("does not coalesce disjoint invalidations", async () => {
    const engine = new FakeEngine(new Map());
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    const changes: CellChange[] = [
      { sheet: "Sheet1", address: "A1", value: 1 },
      { sheet: "Sheet1", address: "C1", value: 2 }
    ];
    provider.applyRecalcChanges(changes);
    await flushMicrotasks();

    expect(updates).toHaveLength(2);
    const ranges = updates.map((u) => (u.type === "cells" ? u.range : null)).filter(Boolean);
    expect(ranges).toContainEqual({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    expect(ranges).toContainEqual({ startRow: 0, endRow: 1, startCol: 2, endCol: 3 });
  });

  it("falls back to a bounding-box invalidation for large change sets", async () => {
    const engine = new FakeEngine(new Map());
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10_000, colCount: 10_000 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    const changes: CellChange[] = [];
    for (let i = 0; i < 300; i++) {
      const col0 = i * 2; // ensure gaps so the regular coalescer would keep many ranges
      changes.push({ sheet: "Sheet1", address: toA1(0, col0), value: i });
    }

    provider.applyRecalcChanges(changes);
    await flushMicrotasks();

    expect(updates).toEqual([
      {
        type: "cells",
        range: { startRow: 0, endRow: 1, startCol: 0, endCol: 599 }
      }
    ]);
  });

  it("can recalculate via engine and update cache + subscribers", async () => {
    const changes: CellChange[] = [
      { sheet: "Sheet1", address: "A1", value: 1 },
      { sheet: "Sheet1", address: "B1", value: 2 }
    ];

    const engine = new FakeEngine(new Map(), changes);
    const cache = new EngineCellCache(engine as unknown as EngineClient);
    const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

    const updates: CellProviderUpdate[] = [];
    provider.subscribe((update) => updates.push(update));

    await provider.recalculate();
    await flushMicrotasks();

    expect(cache.getValue(0, 0)).toBe(1);
    expect(cache.getValue(0, 1)).toBe(2);
    expect(updates).toEqual([{ type: "cells", range: { startRow: 0, endRow: 1, startCol: 0, endCol: 2 } }]);
  });

  it("does not emit unhandled rejections when recalculate is fire-and-forgotten and fails", async () => {
    const unhandled: unknown[] = [];
    const handler = (reason: unknown) => unhandled.push(reason);
    process.on("unhandledRejection", handler);

    try {
      const engine = {
        async recalculate() {
          throw new Error("recalc failed");
        },
      } as any;
      const cache = new EngineCellCache(engine as unknown as EngineClient);
      const provider = new EngineGridProvider({ cache, rowCount: 10, colCount: 10 });

      void provider.recalculate();

      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(unhandled).toEqual([]);
    } finally {
      process.off("unhandledRejection", handler);
    }
  });
});
