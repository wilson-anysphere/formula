import { describe, expect, it } from "vitest";

import { WorkbookSheetStore, generateDefaultSheetName } from "../workbookSheetStore";

describe("WorkbookSheetStore", () => {
  it("enforces Excel-like sheet name validation", () => {
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    expect(() => store.rename("s1", "   ")).toThrow(/empty/i);
    expect(() => store.rename("s1", "A".repeat(32))).toThrow(/31/);
    expect(() => store.rename("s1", "Bad:Name")).toThrow(/cannot contain/i);
    expect(() => store.rename("s2", "sheet1")).toThrow(/duplicate/i);
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

    const store2 = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    store2.hide("s2");
    expect(store2.getById("s2")?.visibility).toBe("hidden");
    expect(() => store2.hide("s1")).toThrow(/last visible/i);

    store2.unhide("s2");
    expect(store2.getById("s2")?.visibility).toBe("visible");
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
});

