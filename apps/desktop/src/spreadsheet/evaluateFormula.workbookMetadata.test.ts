import { describe, expect, it } from "vitest";

import { evaluateFormula } from "./evaluateFormula.js";

describe("evaluateFormula workbook metadata functions", () => {
  it('returns "" for CELL("filename") when workbook metadata is missing', () => {
    const value = evaluateFormula('=CELL("filename")', () => null, {
      workbookFileMetadata: { directory: null, filename: null },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("");
  });

  it('returns the workbook directory for INFO("directory") when metadata is present', () => {
    const value = evaluateFormula('=INFO("directory")', () => null, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
    });
    expect(value).toBe("/tmp/");
  });

  it('returns #N/A for INFO("directory") when workbook is unsaved', () => {
    const value = evaluateFormula('=INFO("directory")', () => null, {
      workbookFileMetadata: { directory: null, filename: null },
    });
    expect(value).toBe("#N/A");
  });

  it('returns #N/A for INFO("directory") when only filename is known (web-style metadata)', () => {
    const value = evaluateFormula('=INFO("directory")', () => null, {
      workbookFileMetadata: { directory: null, filename: "book.xlsx" },
    });
    expect(value).toBe("#N/A");
  });

  it('formats CELL("filename") as dir + [filename] + sheet', () => {
    const value = evaluateFormula('=CELL("filename")', () => null, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("/tmp/[book.xlsx]Sheet1");
  });

  it('uses the reference sheet for CELL("filename", reference)', () => {
    const getCellValue = (_addr: string) => null;
    const value = evaluateFormula('=CELL("filename",Sheet2!A1)', getCellValue, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("/tmp/[book.xlsx]Sheet2");
  });

  it("falls back to currentSheetName for unqualified references", () => {
    const value = evaluateFormula('=CELL("filename",A1)', () => null, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("/tmp/[book.xlsx]Sheet1");
  });

  it("uses currentSheetName even when cellAddress uses an internal sheet id", () => {
    const value = evaluateFormula('=CELL("filename",A1)', () => null, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
      cellAddress: "sheet_123!B2",
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("/tmp/[book.xlsx]Sheet1");
  });

  it("infers a trailing path separator when directory is missing one", () => {
    const value = evaluateFormula('=CELL("filename")', () => null, {
      workbookFileMetadata: { directory: "C:\\tmp", filename: "book.xlsx" },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("C:\\tmp\\[book.xlsx]Sheet1");

    const dir = evaluateFormula('=INFO("directory")', () => null, {
      workbookFileMetadata: { directory: "C:\\tmp", filename: "book.xlsx" },
    });
    expect(dir).toBe("C:\\tmp\\");
  });

  it("trims workbook metadata strings (defensive)", () => {
    const value = evaluateFormula('=CELL("filename")', () => null, {
      workbookFileMetadata: { directory: "  /tmp  ", filename: "  book.xlsx  " },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("/tmp/[book.xlsx]Sheet1");

    const dir = evaluateFormula('=INFO("directory")', () => null, {
      workbookFileMetadata: { directory: "  /tmp  ", filename: "  book.xlsx  " },
    });
    expect(dir).toBe("/tmp/");
  });

  it("supports nested usage inside string concatenation", () => {
    const value = evaluateFormula('="file="&CELL("filename")', () => null, {
      workbookFileMetadata: { directory: "/tmp/", filename: "book.xlsx" },
      currentSheetName: "Sheet1",
    });
    expect(value).toBe("file=/tmp/[book.xlsx]Sheet1");
  });
});
