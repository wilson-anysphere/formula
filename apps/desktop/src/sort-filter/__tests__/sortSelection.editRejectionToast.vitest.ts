/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { sortSelection } from "../sortSelection.js";

describe("sortSelection edit rejection toasts", () => {
  beforeEach(() => {
    document.body.innerHTML = `<div id="toast-root"></div>`;
  });

  it("aborts sorting when any cell is not writable (permission)", () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", { row: 0, col: 0 }, [[{ value: 2 }], [{ value: 1 }]]);
    // Simulate a protected cell in the selection.
    (doc as any).canEditCell = ({ row }: { sheetId: string; row: number; col: number }) => row !== 1;

    const app = {
      isReadOnly: () => false,
      getSelectionRanges: () => [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }],
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getDocument: () => doc,
      getCellComputedValueForSheet: (sheetId: string, cell: { row: number; col: number }) =>
        (doc.getCell(sheetId, cell) as any)?.value ?? null,
      focus: () => {},
    } as any;

    sortSelection(app, { order: "ascending" });

    // No document changes should occur.
    expect(doc.getCell("Sheet1", { row: 0, col: 0 })).toMatchObject({ value: 2 });
    expect(doc.getCell("Sheet1", { row: 1, col: 0 })).toMatchObject({ value: 1 });

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("sort");
    expect(content).toContain("Read-only");
  });

  it("aborts sorting when any cell is not writable (missing encryption key)", () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", { row: 0, col: 0 }, [[{ value: 2 }], [{ value: 1 }]]);
    (doc as any).canEditCell = () => false;

    const app = {
      isReadOnly: () => false,
      getSelectionRanges: () => [{ startRow: 0, endRow: 1, startCol: 0, endCol: 0 }],
      getActiveCell: () => ({ row: 0, col: 0 }),
      getCurrentSheetId: () => "Sheet1",
      getDocument: () => doc,
      getCellComputedValueForSheet: (sheetId: string, cell: { row: number; col: number }) =>
        (doc.getCell(sheetId, cell) as any)?.value ?? null,
      inferCollabEditRejection: () => ({ rejectionReason: "encryption" as const }),
      focus: () => {},
    } as any;

    sortSelection(app, { order: "ascending" });

    expect(doc.getCell("Sheet1", { row: 0, col: 0 })).toMatchObject({ value: 2 });
    expect(doc.getCell("Sheet1", { row: 1, col: 0 })).toMatchObject({ value: 1 });

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Missing encryption key");
  });
});

