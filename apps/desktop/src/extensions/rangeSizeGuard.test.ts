import { describe, expect, it } from "vitest";

import {
  DEFAULT_EXTENSION_RANGE_CELL_LIMIT,
  assertExtensionRangeWithinLimits,
  getExtensionRangeSize,
  normalizeExtensionRange,
} from "./rangeSizeGuard";

describe("extension range size guard", () => {
  it("normalizes inverted ranges", () => {
    expect(normalizeExtensionRange({ startRow: 10, endRow: 5, startCol: 7, endCol: 3 })).toEqual({
      startRow: 5,
      endRow: 10,
      startCol: 3,
      endCol: 7,
    });
  });

  it("computes cell count for inclusive ranges", () => {
    expect(getExtensionRangeSize({ startRow: 0, endRow: 9, startCol: 0, endCol: 9 })).toEqual({
      rows: 10,
      cols: 10,
      cellCount: 100,
    });
  });

  it("allows ranges up to the configured limit", () => {
    expect(() =>
      assertExtensionRangeWithinLimits({ startRow: 0, endRow: 0, startCol: 0, endCol: DEFAULT_EXTENSION_RANGE_CELL_LIMIT - 1 }),
    ).not.toThrow();
  });

  it("throws when the range exceeds the configured limit", () => {
    expect(() => assertExtensionRangeWithinLimits({ startRow: 0, endRow: 999, startCol: 0, endCol: 999 })).toThrow(/too large/i);
  });
});

