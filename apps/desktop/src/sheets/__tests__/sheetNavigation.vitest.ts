import { describe, expect, it } from "vitest";

import { pickAdjacentVisibleSheetId } from "../sheetNavigation";

describe("pickAdjacentVisibleSheetId", () => {
  it("prefers the next visible sheet to the right, otherwise the previous visible sheet", () => {
    const sheets = [
      { id: "s1", visibility: "visible" as const },
      { id: "s2", visibility: "visible" as const },
      { id: "s3", visibility: "visible" as const },
    ];

    expect(pickAdjacentVisibleSheetId(sheets, "s2")).toBe("s3");
    expect(pickAdjacentVisibleSheetId(sheets, "s3")).toBe("s2");
    expect(pickAdjacentVisibleSheetId(sheets, "s1")).toBe("s2");
  });

  it("skips hidden sheets when searching to the right/left", () => {
    const sheets = [
      { id: "s1", visibility: "visible" as const },
      { id: "s2", visibility: "hidden" as const },
      { id: "s3", visibility: "veryHidden" as const },
      { id: "s4", visibility: "visible" as const },
    ];

    // From s1, the next visible sheet is s4 (skip hidden/veryHidden).
    expect(pickAdjacentVisibleSheetId(sheets, "s1")).toBe("s4");
    // From s4, fall back to the previous visible sheet (s1).
    expect(pickAdjacentVisibleSheetId(sheets, "s4")).toBe("s1");
  });

  it("handles reference sheets that are themselves hidden (still chooses nearest visible)", () => {
    const sheets = [
      { id: "s1", visibility: "visible" as const },
      { id: "s2", visibility: "hidden" as const },
      { id: "s3", visibility: "visible" as const },
    ];

    // Next visible to the right of hidden s2 is s3.
    expect(pickAdjacentVisibleSheetId(sheets, "s2")).toBe("s3");
  });

  it("returns null when the reference sheet does not exist", () => {
    const sheets = [{ id: "s1", visibility: "visible" as const }];
    expect(pickAdjacentVisibleSheetId(sheets, "missing")).toBeNull();
  });
});

