import { describe, expect, it } from "vitest";
import { DEFAULT_GRID_THEME } from "../../theme/GridTheme";
import { formatCellDisplayText, resolveCellTextColor, resolveCellTextColorWithTheme } from "../CanvasGridRenderer";

describe("CanvasGridRenderer cell value formatting", () => {
  it("renders booleans as Excel-style TRUE/FALSE strings", () => {
    expect(formatCellDisplayText(true)).toBe("TRUE");
    expect(formatCellDisplayText(false)).toBe("FALSE");
  });

  it("defaults error strings (#...) to red text unless a color is explicitly set", () => {
    expect(resolveCellTextColor("#DIV/0!", undefined)).toBe(DEFAULT_GRID_THEME.errorText);
    expect(resolveCellTextColor("#NAME?", undefined)).toBe(DEFAULT_GRID_THEME.errorText);
    expect(resolveCellTextColor("#DIV/0!", "#00ff00")).toBe("#00ff00");
    expect(resolveCellTextColor("hello", undefined)).toBe(DEFAULT_GRID_THEME.cellText);
    expect(resolveCellTextColor(true, undefined)).toBe(DEFAULT_GRID_THEME.cellText);

    expect(resolveCellTextColorWithTheme("#DIV/0!", undefined, { cellText: "rebeccapurple", errorText: "hotpink" })).toBe(
      "hotpink"
    );
    expect(resolveCellTextColorWithTheme("hello", undefined, { cellText: "rebeccapurple", errorText: "hotpink" })).toBe(
      "rebeccapurple"
    );
  });
});
