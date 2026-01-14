// @vitest-environment jsdom

import { act } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { handleCustomSortCommand } from "../openCustomSortDialog.js";
import * as sortSelection from "../sortSelection.js";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

describe("custom sort command wiring", () => {
  beforeEach(() => {
    document.body.replaceChildren();
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);
  });

  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it("mounts the sort dialog and applies the default sort spec", async () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", "A1:B2", [
      ["Name", "Age"],
      ["Alice", 30],
    ]);

    const applySpy = vi.spyOn(sortSelection, "applySortSpecToSelection").mockReturnValue(true);

    let handled = false;
    await act(async () => {
      handled = handleCustomSortCommand("home.editing.sortFilter.customSort", {
        isEditing: () => false,
        getDocument: () => doc,
        getSheetId: () => "Sheet1",
        getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        getCellValue: (sheetId, cell) => {
          const state = doc.getCell(sheetId, cell) as { value: any };
          return (state?.value ?? null) as any;
        },
        focusGrid: () => {},
      });
    });

    expect(handled).toBe(true);

    const dialog = document.querySelector<HTMLDialogElement>("dialog.custom-sort-dialog");
    expect(dialog).not.toBeNull();

    // The dialog should render the SortDialog component content.
    expect(dialog?.querySelector('[data-testid="sort-dialog"]')).not.toBeNull();

    const okBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="sort-dialog-ok"]');
    expect(okBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      okBtn!.click();
    });

    expect(applySpy).toHaveBeenCalledTimes(1);
    const call = applySpy.mock.calls[0]?.[0] as any;
    expect(call?.sheetId).toBe("Sheet1");
    expect(call?.selection).toEqual({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
    expect(call?.spec).toEqual({ keys: [{ column: 0, order: "ascending" }], hasHeader: true });

    // OK should close + clean up the dialog.
    expect(document.querySelector("dialog.custom-sort-dialog")).toBeNull();
  });

  it("switches to generic column labels when headers are disabled", async () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", "A1:B2", [
      ["Name", "Age"],
      ["Alice", 30],
    ]);

    const applySpy = vi.spyOn(sortSelection, "applySortSpecToSelection").mockReturnValue(true);

    await act(async () => {
      handleCustomSortCommand("data.sortFilter.sort.customSort", {
        isEditing: () => false,
        getDocument: () => doc,
        getSheetId: () => "Sheet1",
        getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        getCellValue: (sheetId, cell) => {
          const state = doc.getCell(sheetId, cell) as { value: any };
          return (state?.value ?? null) as any;
        },
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.custom-sort-dialog");
    expect(dialog).not.toBeNull();

    const headerToggle = dialog!.querySelector<HTMLInputElement>('[data-testid="sort-dialog-has-header"]');
    expect(headerToggle).toBeInstanceOf(HTMLInputElement);
    expect(headerToggle?.checked).toBe(true);

    const columnSelect = dialog!.querySelector<HTMLSelectElement>('[data-testid="sort-dialog-column-0"]');
    expect(columnSelect).toBeInstanceOf(HTMLSelectElement);
    const headerOptions = Array.from(columnSelect!.querySelectorAll("option")).map((o) => o.textContent);
    expect(headerOptions).toEqual(["Name", "Age"]);

    await act(async () => {
      headerToggle!.click();
    });

    expect(headerToggle?.checked).toBe(false);
    const fallbackOptions = Array.from(columnSelect!.querySelectorAll("option")).map((o) => o.textContent);
    expect(fallbackOptions).toEqual(["A", "B"]);

    await act(async () => {
      dialog!.querySelector<HTMLButtonElement>('[data-testid="sort-dialog-ok"]')!.click();
    });

    expect(applySpy).toHaveBeenCalledTimes(1);
    expect((applySpy.mock.calls[0]?.[0] as any)?.spec?.hasHeader).toBe(false);
  });

  it("blocks custom sort in read-only mode (shows toast + does not open dialog)", async () => {
    const focusGrid = vi.fn();

    let handled = false;
    await act(async () => {
      handled = handleCustomSortCommand("data.sortFilter.sort.customSort", {
        isEditing: () => false,
        isReadOnly: () => true,
        // These should never be consulted when read-only.
        getDocument: () => {
          throw new Error("getDocument should not be called");
        },
        getSheetId: () => {
          throw new Error("getSheetId should not be called");
        },
        getSelectionRanges: () => {
          throw new Error("getSelectionRanges should not be called");
        },
        getCellValue: () => {
          throw new Error("getCellValue should not be called");
        },
        focusGrid,
      });
    });

    expect(handled).toBe(true);
    expect(focusGrid).toHaveBeenCalledTimes(1);
    expect(document.querySelector("dialog.custom-sort-dialog")).toBeNull();
    const toast = document.querySelector<HTMLElement>('[data-testid="toast"]');
    expect(toast?.textContent).toMatch(/Read-only/i);
    // Clear the toast timeout so the test environment can tear down cleanly.
    toast?.click();
  });

  it("blocks custom sort when the selection contains a non-writable cell (encryption-aware toast)", async () => {
    const doc = new DocumentController();
    doc.setRangeValues("Sheet1", "A1:B2", [
      ["Name", "Age"],
      ["Alice", 30],
    ]);

    // Simulate a collab/protected/encrypted workbook where writes are blocked.
    (doc as any).canEditCell = () => false;

    const applySpy = vi.spyOn(sortSelection, "applySortSpecToSelection").mockReturnValue(true);

    await act(async () => {
      handleCustomSortCommand("data.sortFilter.sort.customSort", {
        isEditing: () => false,
        getDocument: () => doc,
        getSheetId: () => "Sheet1",
        getSelectionRanges: () => [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        getCellValue: (sheetId, cell) => {
          const state = doc.getCell(sheetId, cell) as { value: any };
          return (state?.value ?? null) as any;
        },
        inferCollabEditRejection: () => ({ rejectionReason: "encryption" }),
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.custom-sort-dialog");
    expect(dialog).not.toBeNull();

    const okBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="sort-dialog-ok"]');
    expect(okBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      okBtn!.click();
    });

    // Should not attempt to apply the sort at all.
    expect(applySpy).not.toHaveBeenCalled();
    // Dialog remains open so the user can adjust selection/spec.
    expect(document.querySelector("dialog.custom-sort-dialog")).not.toBeNull();

    const toast = document.querySelector<HTMLElement>('[data-testid="toast"]');
    expect(toast?.textContent).toMatch(/Missing encryption key/i);
    toast?.click();
  });
});
