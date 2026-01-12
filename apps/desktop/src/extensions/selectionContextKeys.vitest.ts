import { describe, expect, it } from "vitest";

import { addCellToSelection, createSelection, selectRows } from "../selection/selection";
import type { GridLimits } from "../selection/types";
import { deriveSelectionContextKeys } from "./selectionContextKeys.js";
import { evaluateWhenClause } from "./whenClause.js";

const limits: GridLimits = { maxRows: 100, maxCols: 50 };

describe("selection context keys", () => {
  it("supports common menu when-clauses", () => {
    const start = createSelection({ row: 2, col: 2 }, limits);
    const selection = selectRows(start, 3, 5, {}, limits);

    const keys = {
      sheetName: "Sheet1",
      cellHasValue: true,
      ...deriveSelectionContextKeys(selection),
    };

    const lookup = (key: string) => (keys as any)[key];

    expect(evaluateWhenClause("selectionType == 'row'", lookup)).toBe(true);
    expect(evaluateWhenClause("!isSingleCell", lookup)).toBe(true);
    expect(evaluateWhenClause("hasSelection && cellHasValue", lookup)).toBe(true);
  });

  it("treats multi-range selections as hasSelection", () => {
    const start = createSelection({ row: 0, col: 0 }, limits);
    const selection = addCellToSelection(start, { row: 4, col: 4 }, limits);
    const keys = deriveSelectionContextKeys(selection);

    expect(keys.selectionType).toBe("multi");
    expect(keys.hasSelection).toBe(true);
    expect(keys.isMultiRange).toBe(true);

    const lookup = (key: string) => (keys as any)[key];
    expect(evaluateWhenClause("hasSelection", lookup)).toBe(true);
  });
});

