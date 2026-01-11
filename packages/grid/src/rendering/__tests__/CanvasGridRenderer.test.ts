import { describe, expect, it } from "vitest";
import { formatCellDisplayText } from "../CanvasGridRenderer";

describe("CanvasGridRenderer cell value formatting", () => {
  it("renders booleans as Excel-style TRUE/FALSE strings", () => {
    expect(formatCellDisplayText(true)).toBe("TRUE");
    expect(formatCellDisplayText(false)).toBe("FALSE");
  });
});

