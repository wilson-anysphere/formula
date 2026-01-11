import { describe, expect, it, vi } from "vitest";
import { EngineCellProvider, type EngineClientLike } from "./EngineCellProvider";

describe("EngineCellProvider", () => {
  it("renders headers and maps prefetch ranges to A1 ranges", async () => {
    const engine: EngineClientLike = {
      getRange: vi.fn(async (range: string) => {
        expect(range).toBe("A1:B1");
        return [[{ value: 1 }, { value: 3 }]];
      })
    };

    const provider = new EngineCellProvider({ engine, rowCount: 100, colCount: 100 });

    expect(provider.getCell(0, 0)?.value).toBeNull();
    expect(provider.getCell(0, 1)?.value).toBe("A");
    expect(provider.getCell(0, 2)?.value).toBe("B");
    expect(provider.getCell(1, 0)?.value).toBe(1);

    const updates: unknown[] = [];
    const unsubscribe = provider.subscribe((update) => updates.push(update));

    await provider.prefetch({ startRow: 0, endRow: 2, startCol: 0, endCol: 3 });

    unsubscribe();

    expect(provider.getCell(1, 1)?.value).toBe(1);
    expect(provider.getCell(1, 2)?.value).toBe(3);

    expect(updates).toEqual([
      {
        type: "cells",
        range: { startRow: 1, endRow: 2, startCol: 1, endCol: 3 }
      }
    ]);
  });

  it("converts column indices to Excel-style letters", () => {
    const engine: EngineClientLike = { getRange: vi.fn(async () => [[]]) };
    const provider = new EngineCellProvider({ engine, rowCount: 10, colCount: 60 });

    // grid col 27 => engine col 26 => AA
    expect(provider.getCell(0, 27)?.value).toBe("AA");
  });
});

