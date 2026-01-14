import { describe, expect, it, vi } from "vitest";

import { DEFAULT_GRID_FONT_FAMILY } from "@formula/grid/node";
import { resolveCssVar } from "../../../theme/cssVars.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

type CellState = { value: unknown; formula: string | null; styleId?: number };

function createProvider(options: {
  getSheetId: () => string;
  getCell: (sheetId: string, coord: { row: number; col: number }) => CellState | null;
  headerRows?: number;
  headerCols?: number;
}) {
  const headerRows = options.headerRows ?? 1;
  const headerCols = options.headerCols ?? 1;
  const doc = {
    getCell: vi.fn(options.getCell),
    on: vi.fn(() => () => {})
  };

  const provider = new DocumentCellProvider({
    document: doc as any,
    getSheetId: options.getSheetId,
    headerRows,
    headerCols,
    rowCount: headerRows + 10,
    colCount: headerCols + 10,
    showFormulas: () => false,
    getComputedValue: () => null
  });

  return { provider, doc };
}

describe("DocumentCellProvider typography defaults (shared grid)", () => {
  it("uses system font for header cells while allowing data cells to omit fontFamily (renderer default)", () => {
    const { provider, doc } = createProvider({
      getSheetId: () => "sheet-1",
      getCell: () => ({ value: "hello", formula: null, styleId: 0 })
    });

    const colHeader = provider.getCell(0, 5);
    expect(colHeader?.value).toBe("E");
    expect(colHeader?.style?.fontFamily).toBe(resolveCssVar("--font-sans", { fallback: DEFAULT_GRID_FONT_FAMILY }));

    const rowHeader = provider.getCell(5, 0);
    expect(rowHeader?.value).toBe(5);
    expect(rowHeader?.style?.fontFamily).toBe(resolveCssVar("--font-sans", { fallback: DEFAULT_GRID_FONT_FAMILY }));

    const dataCell = provider.getCell(1, 1);
    expect(dataCell?.value).toBe("hello");
    expect(dataCell?.style?.fontFamily).toBeUndefined();

    // Sanity: header cells are synthetic (DocumentController shouldn't be queried).
    expect(doc.getCell).toHaveBeenCalledTimes(1);
  });
});
