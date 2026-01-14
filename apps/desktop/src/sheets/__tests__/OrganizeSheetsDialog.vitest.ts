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
  for (const dialog of Array.from(document.querySelectorAll("dialog"))) {
    try {
      (dialog as any).close?.();
    } catch {
      dialog.remove();
    }
  }
  document.body.innerHTML = "";
  delete (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
  vi.restoreAllMocks();
});

describe("OrganizeSheetsDialog", () => {
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
});
