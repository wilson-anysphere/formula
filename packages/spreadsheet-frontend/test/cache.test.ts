import { describe, expect, it } from "vitest";
import type { CellChange, CellData as EngineCellData, CellScalar, EngineClient } from "@formula/engine";
import { EngineCellCache, fromA1, toA1 } from "../src/index.js";

class FakeEngine {
  calls: Array<{ range: string; sheet?: string }> = [];

  constructor(private readonly values: Map<string, CellScalar>) {}

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

  async recalculate(): Promise<CellChange[]> {
    return [];
  }
}

describe("EngineCellCache", () => {
  it("evicts oldest entries when exceeding maxEntries", async () => {
    const values = new Map<string, CellScalar>([
      ["Sheet1!A1", 1],
      ["Sheet1!B1", 2],
      ["Sheet1!C1", 3],
      ["Sheet1!D1", 4]
    ]);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient, { maxEntries: 3 });

    await cache.prefetch({ startRow0: 0, endRow0Exclusive: 1, startCol0: 0, endCol0Exclusive: 4 });

    // Oldest entry (A1) should be evicted.
    expect(cache.hasValue(0, 0)).toBe(false);
    expect(cache.getValue(0, 1)).toBe(2);
    expect(cache.getValue(0, 2)).toBe(3);
    expect(cache.getValue(0, 3)).toBe(4);
  });

  it("refreshes insertion order on updates (LRU-by-prefetch)", async () => {
    const values = new Map<string, CellScalar>([
      ["Sheet1!A1", 1],
      ["Sheet1!B1", 2],
      ["Sheet1!C1", 3],
      ["Sheet1!D1", 4]
    ]);

    const engine = new FakeEngine(values);
    const cache = new EngineCellCache(engine as unknown as EngineClient, { maxEntries: 3 });

    await cache.prefetch({ startRow0: 0, endRow0Exclusive: 1, startCol0: 0, endCol0Exclusive: 3 });
    cache.applyRecalcChanges([{ sheet: "Sheet1", address: "A1", value: 10 }]);
    await cache.prefetch({ startRow0: 0, endRow0Exclusive: 1, startCol0: 3, endCol0Exclusive: 4 });

    // Updating A1 should have kept it from being evicted when adding D1.
    expect(cache.hasValue(0, 1)).toBe(false); // B1 is now the oldest
    expect(cache.getValue(0, 0)).toBe(10);
    expect(cache.getValue(0, 2)).toBe(3);
    expect(cache.getValue(0, 3)).toBe(4);
  });
});
