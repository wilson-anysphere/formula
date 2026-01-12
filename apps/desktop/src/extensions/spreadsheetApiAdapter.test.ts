import { describe, expect, it } from "vitest";

import { buildSheetNameToIdMap, parseSheetQualifiedA1Range, resolveSheetIdByName } from "./spreadsheetApiAdapter";

describe("desktop spreadsheetApi adapter helpers", () => {
  describe("parseSheetQualifiedA1Range", () => {
    it("parses unqualified single-cell refs", () => {
      expect(parseSheetQualifiedA1Range("A1")).toEqual({
        sheetName: null,
        ref: "A1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      });
    });

    it("parses unqualified range refs", () => {
      expect(parseSheetQualifiedA1Range("B2:C3")).toEqual({
        sheetName: null,
        ref: "B2:C3",
        startRow: 1,
        startCol: 1,
        endRow: 2,
        endCol: 2,
      });
    });

    it("parses sheet-qualified refs", () => {
      expect(parseSheetQualifiedA1Range("Sheet1!A1:B2")).toEqual({
        sheetName: "Sheet1",
        ref: "A1:B2",
        startRow: 0,
        startCol: 0,
        endRow: 1,
        endCol: 1,
      });
    });

    it("parses quoted sheet-qualified refs", () => {
      expect(parseSheetQualifiedA1Range("'My Sheet'!A1")).toEqual({
        sheetName: "My Sheet",
        ref: "A1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
      });
    });
  });

  describe("buildSheetNameToIdMap / resolveSheetIdByName", () => {
    it("maps display names to ids", () => {
      const workbookSheetNames = new Map([
        ["sheet-1", "Sheet1"],
        ["sheet-2", "Sheet2"],
      ]);
      const sheetIds = ["sheet-1", "sheet-2"];
      const map = buildSheetNameToIdMap(sheetIds, workbookSheetNames);
      expect(map.get("Sheet1")).toBe("sheet-1");
      expect(map.get("Sheet2")).toBe("sheet-2");
    });

    it("falls back to ids when display names are missing", () => {
      const workbookSheetNames = new Map<string, string>();
      const sheetIds = ["Sheet1"];
      const map = buildSheetNameToIdMap(sheetIds, workbookSheetNames);
      expect(map.get("Sheet1")).toBe("Sheet1");
    });

    it("throws when resolving unknown sheet names", () => {
      const workbookSheetNames = new Map([["sheet-1", "Sheet1"]]);
      expect(() =>
        resolveSheetIdByName({ sheetName: "Missing", sheetIds: ["sheet-1"], workbookSheetNames }),
      ).toThrow(/Unknown sheet/i);
    });

    it("throws when sheet names are ambiguous", () => {
      const workbookSheetNames = new Map([
        ["sheet-1", "Data"],
        ["sheet-2", "Data"],
      ]);
      expect(() => buildSheetNameToIdMap(["sheet-1", "sheet-2"], workbookSheetNames)).toThrow(/Duplicate sheet name/i);
    });
  });
});

