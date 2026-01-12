import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentCellProvider } from "../documentCellProvider.js";

describe("DocumentCellProvider.prefetch", () => {
  it("does not synchronously warm via getCell()", () => {
    const doc = new DocumentController();
    const provider = new DocumentCellProvider({
      document: doc,
      getSheetId: () => "Sheet1",
      headerRows: 1,
      headerCols: 1,
      rowCount: 100,
      colCount: 100,
      showFormulas: () => false,
      getComputedValue: () => null,
    });

    const spy = vi.spyOn(provider, "getCell");

    provider.prefetch({ startRow: 0, endRow: 50, startCol: 0, endCol: 50 });

    expect(spy).not.toHaveBeenCalled();
  });
});
