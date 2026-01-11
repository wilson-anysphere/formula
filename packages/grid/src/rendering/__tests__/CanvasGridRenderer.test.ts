import { describe, expect, it } from "vitest";
import { formatCellDisplayText, resolveCellTextColor } from "../CanvasGridRenderer";

describe("CanvasGridRenderer cell value formatting", () => {
  it("renders booleans as Excel-style TRUE/FALSE strings", () => {
    expect(formatCellDisplayText(true)).toBe("TRUE");
    expect(formatCellDisplayText(false)).toBe("FALSE");
  });

  it("defaults error strings (#...) to red text unless a color is explicitly set", () => {
    expect(resolveCellTextColor("#DIV/0!", undefined)).toBe("#cc0000");
    expect(resolveCellTextColor("#NAME?", undefined)).toBe("#cc0000");
    expect(resolveCellTextColor("#DIV/0!", "#00ff00")).toBe("#00ff00");
    expect(resolveCellTextColor("hello", undefined)).toBe("#111111");
    expect(resolveCellTextColor(true, undefined)).toBe("#111111");
  });
});
