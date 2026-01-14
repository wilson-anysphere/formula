// @vitest-environment jsdom
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { openOrganizeSheetsDialog } from "../OrganizeSheetsDialog";
import { rewriteDocumentFormulasForSheetRename } from "../sheetFormulaRewrite";
import { WorkbookSheetStore } from "../workbookSheetStore";

function ensureDialogPolyfill(): void {
  // JSDOM doesn't implement <dialog>. Patch in a minimal `showModal` / `close`
  // so `openOrganizeSheetsDialog` can be exercised in unit tests.
  const proto = (globalThis as any).HTMLDialogElement?.prototype;
  if (!proto) return;

  proto.showModal ??= function showModal(this: HTMLDialogElement) {
    this.setAttribute("open", "");
  };
  proto.close ??= function close(this: HTMLDialogElement, returnValue?: string) {
    (this as any).returnValue = returnValue ?? "";
    this.removeAttribute("open");
    this.dispatchEvent(new Event("close"));
  };
}

beforeEach(() => {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  document.body.innerHTML = "";
  ensureDialogPolyfill();
});

afterEach(() => {
  // Best-effort cleanup: close any dialogs so React roots unmount.
  //
  // Wrap in `act` to avoid React 18 warnings about untracked state updates triggered
  // by `root.unmount()` in the dialog close handler.
  act(() => {
    for (const dialog of Array.from(document.querySelectorAll("dialog"))) {
      try {
        (dialog as any).close?.();
      } catch {
        dialog.remove();
      }
    }
  });
  document.body.innerHTML = "";
  delete (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
  vi.restoreAllMocks();
});

describe("OrganizeSheetsDialog", () => {
  it("does not open while the spreadsheet is editing", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => "s1",
        activateSheet: () => {},
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => true,
        focusGrid: () => {},
      });
    });

    expect(document.querySelector('dialog[data-testid="organize-sheets-dialog"]')).toBeNull();
  });

  it("renders a tab color indicator when present", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible", tabColor: { rgb: "FFFF0000" } },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    expect(dialog!.querySelector('[data-testid="organize-sheet-tab-color-s1"]')).toBeInstanceOf(HTMLElement);
    expect(dialog!.querySelector('[data-testid="organize-sheet-tab-color-s2"]')).toBeNull();
  });

  it("renames a sheet and rewrites formulas", async () => {
    const doc = new DocumentController();
    doc.setCellFormula("S1", { row: 0, col: 0 }, "='Budget'!A1");

    const store = new WorkbookSheetStore([
      { id: "S1", name: "Sheet1", visibility: "visible" },
      { id: "S2", name: "Budget", visibility: "visible" },
    ]);

    let activeSheetId = "S1";

    const renameSheetById = async (sheetId: string, newName: string) => {
      const oldName = store.getName(sheetId) ?? sheetId;
      store.rename(sheetId, newName);
      rewriteDocumentFormulasForSheetRename(doc, oldName, store.getName(sheetId) ?? newName);
    };

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById,
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-S2"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      renameBtn!.click();
    });

    const input = dialog!.querySelector<HTMLInputElement>('[data-testid="organize-sheet-rename-input-S2"]');
    expect(input).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      (input as HTMLInputElement).value = "Budget2024";
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
      await Promise.resolve();
    });

    expect(store.getName("S2")).toBe("Budget2024");
    expect(doc.getCell("S1", { row: 0, col: 0 }).formula).toBe("=Budget2024!A1");
  });

  it("surfaces rename validation errors and keeps the rename UI open", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Budget", visibility: "visible" },
    ]);

    let activeSheetId = "s1";
    const renameSheetById = vi.fn(async (sheetId: string, newName: string) => {
      // `WorkbookSheetStore.rename` enforces Excel-like validation (including duplicates).
      store.rename(sheetId, newName);
    });

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById,
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s2"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      renameBtn!.click();
    });

    const input = dialog!.querySelector<HTMLInputElement>('[data-testid="organize-sheet-rename-input-s2"]');
    expect(input).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      (input as HTMLInputElement).value = "Sheet1";
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
      await Promise.resolve();
    });

    expect(renameSheetById).toHaveBeenCalledTimes(1);
    expect(renameSheetById).toHaveBeenCalledWith("s2", "Sheet1");
    expect(store.getName("s2")).toBe("Budget");

    const error = dialog!.querySelector('[data-testid="organize-sheets-error"]');
    expect(error).toBeInstanceOf(HTMLElement);
    expect(error?.textContent?.trim()).toBeTruthy();

    // Ensure the rename UI is still active after the error (Excel-style).
    expect(dialog!.querySelector('[data-testid="organize-sheet-rename-input-s2"]')).toBeInstanceOf(HTMLInputElement);
  });

  it("enforces the 'cannot hide the last visible sheet' invariant", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const hide2 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s2"]');
    expect(hide2).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      hide2!.click();
    });
    expect(store.getById("s2")?.visibility).toBe("hidden");

    const hide1 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s1"]');
    expect(hide1).toBeInstanceOf(HTMLButtonElement);
    expect(hide1!.disabled).toBe(true);

    // Unhide restores normal hide capability.
    const unhide2 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-unhide-s2"]');
    expect(unhide2).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      unhide2!.click();
    });
    expect(store.getById("s2")?.visibility).toBe("visible");

    const hide1After = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s1"]');
    expect(hide1After).toBeInstanceOf(HTMLButtonElement);
    expect(hide1After!.disabled).toBe(false);
  });

  it("prevents deleting the last visible sheet (even if hidden sheets remain)", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const hide2 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s2"]');
    expect(hide2).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      hide2!.click();
    });
    expect(store.getById("s2")?.visibility).toBe("hidden");

    const deleteVisible = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s1"]');
    expect(deleteVisible).toBeInstanceOf(HTMLButtonElement);
    expect(deleteVisible!.disabled).toBe(true);

    const deleteHidden = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s2"]');
    expect(deleteHidden).toBeInstanceOf(HTMLButtonElement);
    expect(deleteHidden!.disabled).toBe(false);
  });

  it("reorders sheets via move up/down buttons", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
      { id: "s3", name: "Sheet3", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const moveDown = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-move-down-s1"]');
    expect(moveDown).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      moveDown!.click();
    });

    expect(store.listAll().map((s) => s.id)).toEqual(["s2", "s1", "s3"]);
  });

  it("deletes the active sheet and activates a fallback", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const deleteBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s1"]');
    expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      deleteBtn!.click();
      await Promise.resolve();
    });

    const confirmBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-confirm-s1"]');
    expect(confirmBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      confirmBtn!.click();
      await Promise.resolve();
    });

    expect(store.getById("s1")).toBeUndefined();
    expect(store.listAll().map((s) => s.id)).toEqual(["s2"]);
    expect(activeSheetId).toBe("s2");
  });

  it("rewrites formulas when deleting a sheet", async () => {
    const doc = new DocumentController();
    doc.setCellFormula("s1", { row: 0, col: 0 }, "=Budget!A1+1");

    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Budget", visibility: "visible" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const deleteBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s2"]');
    expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      deleteBtn!.click();
      await Promise.resolve();
    });

    const confirmBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-confirm-s2"]');
    expect(confirmBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      confirmBtn!.click();
      await Promise.resolve();
    });

    expect(store.getById("s2")).toBeUndefined();
    expect(doc.getCell("s1", { row: 0, col: 0 }).formula).toBe("=#REF!+1");
  });

  it("activating a hidden sheet unhides it first", async () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Hidden", visibility: "hidden" },
    ]);

    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const activateBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-activate-s2"]');
    expect(activateBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      activateBtn!.click();
      await Promise.resolve();
    });

    expect(store.getById("s2")?.visibility).toBe("visible");
    expect(activeSheetId).toBe("s2");
  });

  it("updates the active indicator when formula:sheet-activated fires", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    expect(dialog!.querySelector('[data-testid="organize-sheet-active-s1"]')).toBeInstanceOf(HTMLElement);
    expect(dialog!.querySelector('[data-testid="organize-sheet-active-s2"]')).toBeNull();

    act(() => {
      activeSheetId = "s2";
      window.dispatchEvent(new CustomEvent("formula:sheet-activated", { detail: { sheetId: "s2" } }));
    });

    expect(dialog!.querySelector('[data-testid="organize-sheet-active-s1"]')).toBeNull();
    expect(dialog!.querySelector('[data-testid="organize-sheet-active-s2"]')).toBeInstanceOf(HTMLElement);
  });

  it("disables other sheet actions while renaming (only Save/Cancel remain active)", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s1"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      renameBtn!.click();
    });

    // Save/Cancel should remain enabled.
    const save = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-save-s1"]');
    expect(save).toBeInstanceOf(HTMLButtonElement);
    expect(save!.disabled).toBe(false);

    const cancel = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-cancel-s1"]');
    expect(cancel).toBeInstanceOf(HTMLButtonElement);
    expect(cancel!.disabled).toBe(false);

    // Other actions should be disabled while the rename UI is open.
    const hide = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s1"]');
    expect(hide).toBeInstanceOf(HTMLButtonElement);
    expect(hide!.disabled).toBe(true);

    const del = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s1"]');
    expect(del).toBeInstanceOf(HTMLButtonElement);
    expect(del!.disabled).toBe(true);

    const move = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-move-down-s1"]');
    expect(move).toBeInstanceOf(HTMLButtonElement);
    expect(move!.disabled).toBe(true);

    const activate = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-activate-s1"]');
    expect(activate).toBeInstanceOf(HTMLButtonElement);
    expect(activate!.disabled).toBe(true);

    // Other rows should also be disabled.
    const otherRename = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s2"]');
    expect(otherRename).toBeInstanceOf(HTMLButtonElement);
    expect(otherRename!.disabled).toBe(true);
  });

  it("restores focus via host.focusGrid when closing the dialog", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "s1";
    const focusGrid = vi.fn();

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid,
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const closeBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheets-close"]');
    expect(closeBtn).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      closeBtn!.click();
    });

    expect(focusGrid).toHaveBeenCalledTimes(1);
  });

  it("closes on Escape when no inline state is active", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    let activeSheetId = "s1";
    const focusGrid = vi.fn();

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid,
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const body = dialog!.querySelector<HTMLElement>('[data-testid="organize-sheets-dialog-body"]');
    expect(body).toBeInstanceOf(HTMLElement);

    act(() => {
      body!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    });

    expect(focusGrid).toHaveBeenCalledTimes(1);
    expect(document.querySelector('dialog[data-testid="organize-sheets-dialog"]')).toBeNull();
  });

  it("Escape cancels inline rename before closing the dialog", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s1"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      renameBtn!.click();
    });

    expect(dialog!.querySelector('[data-testid="organize-sheet-rename-input-s1"]')).toBeInstanceOf(HTMLInputElement);

    const body = dialog!.querySelector<HTMLElement>('[data-testid="organize-sheets-dialog-body"]');
    expect(body).toBeInstanceOf(HTMLElement);
    act(() => {
      body!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    });

    // Dialog remains open, but rename UI should be cancelled.
    expect(document.querySelector('dialog[data-testid="organize-sheets-dialog"]')).toBeInstanceOf(HTMLDialogElement);
    expect(dialog!.querySelector('[data-testid="organize-sheet-rename-input-s1"]')).toBeNull();
  });

  it("re-binds to the latest sheet store when the host replaces it (collab-style)", async () => {
    const doc = new DocumentController();
    const store1 = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    const store2 = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "s1";
    let currentStore: WorkbookSheetStore = store1;
    const getStore = () => currentStore;

    act(() => {
      openOrganizeSheetsDialog({
        store: currentStore,
        getStore,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    // Wait a tick for the dialog's effect to install the metadata-change listener.
    await act(async () => {
      await Promise.resolve();
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);
    expect(dialog!.querySelector('[data-testid="organize-sheet-row-s2"]')).toBeNull();

    act(() => {
      currentStore = store2;
      window.dispatchEvent(new CustomEvent("formula:sheet-metadata-changed"));
    });

    expect(dialog!.querySelector('[data-testid="organize-sheet-row-s2"]')).toBeInstanceOf(HTMLElement);
  });

  it("clears inline rename state when a sheet disappears after store replacement", async () => {
    const doc = new DocumentController();
    const store1 = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    const store2 = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);

    let activeSheetId = "s1";
    let currentStore: WorkbookSheetStore = store1;
    const getStore = () => currentStore;

    act(() => {
      openOrganizeSheetsDialog({
        store: currentStore,
        getStore,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        focusGrid: () => {},
      });
    });

    // Wait a tick for the dialog's effect to install the metadata-change listener.
    await act(async () => {
      await Promise.resolve();
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s2"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);
    act(() => {
      renameBtn!.click();
    });

    expect(dialog!.querySelector('[data-testid="organize-sheet-rename-input-s2"]')).toBeInstanceOf(HTMLInputElement);

    act(() => {
      currentStore = store2;
      window.dispatchEvent(new CustomEvent("formula:sheet-metadata-changed"));
    });

    // s2 row + rename UI should be gone, and remaining actions should be enabled.
    expect(dialog!.querySelector('[data-testid="organize-sheet-row-s2"]')).toBeNull();
    expect(dialog!.querySelector('[data-testid="organize-sheet-rename-input-s2"]')).toBeNull();

    const renameS1 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s1"]');
    expect(renameS1).toBeInstanceOf(HTMLButtonElement);
    expect(renameS1!.disabled).toBe(false);

    const activateS1 = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-activate-s1"]');
    expect(activateS1).toBeInstanceOf(HTMLButtonElement);
    expect(activateS1!.disabled).toBe(false);
  });

  it("disables sheet-structure mutations in readOnly mode", () => {
    const doc = new DocumentController();
    const store = new WorkbookSheetStore([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
      { id: "s3", name: "Hidden", visibility: "hidden" },
    ]);
    let activeSheetId = "s1";

    act(() => {
      openOrganizeSheetsDialog({
        store,
        getActiveSheetId: () => activeSheetId,
        activateSheet: (next) => {
          activeSheetId = next;
        },
        renameSheetById: () => {},
        getDocument: () => doc,
        isEditing: () => false,
        readOnly: true,
        focusGrid: () => {},
      });
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="organize-sheets-dialog"]');
    expect(dialog).toBeInstanceOf(HTMLDialogElement);

    // Visible-sheet activation should remain available.
    const activateVisible = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-activate-s2"]');
    expect(activateVisible).toBeInstanceOf(HTMLButtonElement);
    expect(activateVisible!.disabled).toBe(false);

    // Mutations should be disabled.
    const renameBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-rename-s1"]');
    expect(renameBtn).toBeInstanceOf(HTMLButtonElement);
    expect(renameBtn!.disabled).toBe(true);

    const hideBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-hide-s1"]');
    expect(hideBtn).toBeInstanceOf(HTMLButtonElement);
    expect(hideBtn!.disabled).toBe(true);

    const deleteBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-delete-s1"]');
    expect(deleteBtn).toBeInstanceOf(HTMLButtonElement);
    expect(deleteBtn!.disabled).toBe(true);

    const moveDownBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-move-down-s1"]');
    expect(moveDownBtn).toBeInstanceOf(HTMLButtonElement);
    expect(moveDownBtn!.disabled).toBe(true);

    // Hidden-sheet activation requires unhide, so it should be disabled in read-only.
    const activateHidden = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-activate-s3"]');
    expect(activateHidden).toBeInstanceOf(HTMLButtonElement);
    expect(activateHidden!.disabled).toBe(true);

    const unhideBtn = dialog!.querySelector<HTMLButtonElement>('[data-testid="organize-sheet-unhide-s3"]');
    expect(unhideBtn).toBeInstanceOf(HTMLButtonElement);
    expect(unhideBtn!.disabled).toBe(true);
  });
});
