import { describe, expect, it } from "vitest";

import { WorkbookSheetStore, generateDefaultSheetName } from "../workbookSheetStore";

describe("WorkbookSheetStore", () => {
  it("enforces Excel-like sheet name validation", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    expect(() => store.rename("s1", "   ")).toThrow(/blank/i);
    expect(() => store.rename("s1", "A".repeat(32))).toThrow(/31/);
    expect(() => store.rename("s1", "Bad:Name")).toThrow(/invalid character/i);
    expect(() => store.rename("s2", "sheet1")).toThrow(/already exists/i);

    expect(() => store.rename("s1", "'Budget")).toThrow(/apostrophe/i);
    expect(() => store.rename("s1", "Budget'")).toThrow(/apostrophe/i);

    // 🙂 is outside the BMP (2 UTF-16 code units). Excel's 31-char limit is in UTF-16 code units.
    expect(() => store.rename("s1", `${"a".repeat(29)}🙂`)).not.toThrow();
    expect(() => store.rename("s1", `${"a".repeat(30)}🙂`)).toThrow(/31/);
  });

  it("selects the next available default sheet name (case-insensitive)", () => {
    expect(
      generateDefaultSheetName([
        { name: "Sheet1" },
        { name: "Sheet2" },
        { name: "sheet4" },
      ]),
    ).toBe("Sheet3");

    expect(generateDefaultSheetName([{ name: "Sheet1" }, { name: "Sheet3" }])).toBe("Sheet2");
  });

  it("prevents hiding the last visible sheet and supports unhide", () => {
    const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    expect(() => store.hide("s1")).toThrow(/last visible/i);
    expect(() => store.setVisibility("s1", "veryHidden")).toThrow(/last visible/i);

    const store2 = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    store2.hide("s2");
    expect(store2.getById("s2")?.visibility).toBe("hidden");
    expect(() => store2.hide("s1")).toThrow(/last visible/i);

    store2.unhide("s2");
    expect(store2.getById("s2")?.visibility).toBe("visible");

    store2.setVisibility("s2", "veryHidden");
    expect(store2.getById("s2")?.visibility).toBe("veryHidden");
    expect(store2.listVisible().map((s) => s.id)).toEqual(["s1"]);
  });

  it("prevents deleting the last visible sheet (even if hidden sheets remain)", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Hidden", visibility: "hidden" },
    ]);

    expect(() => store.remove("s1")).toThrow(/last visible/i);

    // Deleting hidden sheets is allowed when at least one visible sheet remains.
    store.remove("s2");
    expect(store.listAll().map((s) => s.id)).toEqual(["s1"]);
  });

  it("reorders sheets with move()", () => {
    const store = new WorkbookSheetStore([
      { id: "a", name: "Sheet1", visibility: "visible" },
      { id: "b", name: "Sheet2", visibility: "visible" },
      { id: "c", name: "Sheet3", visibility: "visible" },
    ]);

    store.move("a", 2);
    expect(store.listAll().map((s) => s.id)).toEqual(["b", "c", "a"]);

    store.move("a", 0);
    expect(store.listAll().map((s) => s.id)).toEqual(["a", "b", "c"]);
  });

  it("resolves ids by name case-insensitively", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Budget", visibility: "visible" },
    ]);

    expect(store.resolveIdByName("sheet1")).toBe("s1");
    expect(store.resolveIdByName(" SHEET1 ")).toBe("s1");
    expect(store.resolveIdByName("budget")).toBe("s2");
    expect(store.resolveIdByName("missing")).toBeUndefined();
  });

  it("normalizes tabColor.rgb to uppercase when setting colors", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible", tabColor: { rgb: "#ff0000" } },
    ]);
    expect(store.getById("s1")?.tabColor?.rgb).toBe("#FF0000");

    store.setTabColor("s1", { rgb: "ff00ff00" });
    expect(store.getById("s1")?.tabColor?.rgb).toBe("FF00FF00");
  });

  it("avoids emitting for no-op tabColor updates", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible", tabColor: { rgb: "#FF0000" } },
    ]);
    let calls = 0;
    store.subscribe(() => {
      calls += 1;
    });

    // No-op (case normalization only).
    store.setTabColor("s1", { rgb: "#ff0000" });
    expect(calls).toBe(0);

    store.setTabColor("s1", { rgb: "#00FF00" });
    expect(calls).toBe(1);
  });
});
