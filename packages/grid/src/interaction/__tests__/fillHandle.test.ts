import { describe, expect, it } from "vitest";
import { computeFillPreview, hitTestSelectionHandle } from "../fillHandle";

describe("fillHandle interaction helpers", () => {
  it("computes vertical fill preview (down)", () => {
    const preview = computeFillPreview(
      { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
      { row: 4, col: 1 }
    );
    expect(preview).toEqual({
      axis: "vertical",
      targetRange: { startRow: 3, endRow: 5, startCol: 1, endCol: 2 },
      unionRange: { startRow: 1, endRow: 5, startCol: 1, endCol: 2 }
    });
  });

  it("computes horizontal fill preview (right)", () => {
    const preview = computeFillPreview(
      { startRow: 1, endRow: 2, startCol: 1, endCol: 3 },
      { row: 1, col: 4 }
    );
    expect(preview).toEqual({
      axis: "horizontal",
      targetRange: { startRow: 1, endRow: 2, startCol: 3, endCol: 5 },
      unionRange: { startRow: 1, endRow: 2, startCol: 1, endCol: 5 }
    });
  });

  it("chooses an axis on diagonal drags (prefers vertical on ties)", () => {
    const preview = computeFillPreview(
      { startRow: 10, endRow: 11, startCol: 10, endCol: 11 },
      { row: 13, col: 12 }
    );
    expect(preview?.axis).toBe("vertical");
  });

  it("hit-tests the selection handle", () => {
    const renderer = {
      getFillHandleRect: () => ({ x: 100, y: 50, width: 8, height: 8 })
    } as any;

    expect(hitTestSelectionHandle(renderer, 101, 51)).toBe(true);
    expect(hitTestSelectionHandle(renderer, 10, 10)).toBe(false);
  });
});
