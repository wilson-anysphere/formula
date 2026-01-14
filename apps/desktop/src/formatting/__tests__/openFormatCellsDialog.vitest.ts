// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { DEFAULT_GRID_LIMITS } from "../../selection/selection.js";
import { openFormatCellsDialog } from "../openFormatCellsDialog.js";

function ensureDialogPolyfill(): void {
  // JSDOM doesn't implement <dialog>. Patch in a minimal `showModal` / `close`
  // so `openFormatCellsDialog` can be exercised in unit tests.
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

describe("openFormatCellsDialog formatting performance guards", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);
    ensureDialogPolyfill();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("blocks Apply when selection exceeds the cap and is not a full row/column band", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      // 250k cells -> above DEFAULT_FORMATTING_APPLY_CELL_LIMIT.
      getSelectionRanges: () => [{ startRow: 0, endRow: 499, startCol: 0, endCol: 499 }],
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    // Ensure we have a non-empty patch by toggling Bold.
    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-bold"]')!.click();
    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(spy).toHaveBeenCalledTimes(0);
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Selection is too large to format");

    // Cleanup: remove the toast so its auto-dismiss timeout doesn't keep the test
    // process alive.
    (document.querySelector<HTMLElement>('[data-testid="toast"]') as any)?.click?.();
  });

  it("blocks opening the dialog in read-only mode for non-band selections", () => {
    const doc = new DocumentController();

    openFormatCellsDialog({
      isEditing: () => false,
      isReadOnly: () => true,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focusGrid: () => {},
    });

    expect(document.querySelector("dialog.format-cells-dialog")).toBeNull();
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    // Cleanup: remove the toast so its auto-dismiss timeout doesn't keep the test alive.
    (document.querySelector<HTMLElement>('[data-testid="toast"]') as any)?.click?.();
  });

  it("allows opening the dialog in read-only mode for full row/column band selections", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat").mockImplementation(() => true);

    openFormatCellsDialog({
      isEditing: () => false,
      isReadOnly: () => true,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      // Full column A within legacy limits (10k rows).
      getSelectionRanges: () => [{ startRow: 0, endRow: 9_999, startCol: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-bold"]')!.click();
    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(spy).toHaveBeenCalledTimes(1);
  });

  it("expands full-column selections to Excel bounds before applying", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat").mockImplementation(() => true);

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      // Full column A within legacy limits (10k rows).
      getSelectionRanges: () => [{ startRow: 0, endRow: 9_999, startCol: 0, endCol: 0 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-bold"]')!.click();
    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(spy).toHaveBeenCalledTimes(1);
    const rangeArg = spy.mock.calls[0]?.[1] as any;
    expect(rangeArg?.start?.row).toBe(0);
    expect(rangeArg?.start?.col).toBe(0);
    expect(rangeArg?.end?.row).toBe(DEFAULT_GRID_LIMITS.maxRows - 1);
    expect(rangeArg?.end?.col).toBe(0);
  });

  it("expands full-row selections to Excel bounds before applying", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat").mockImplementation(() => true);

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      // Full row 1 within legacy limits (200 cols).
      getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 199 }],
      getGridLimits: () => ({ maxRows: 10_000, maxCols: 200 }),
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-bold"]')!.click();
    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(spy).toHaveBeenCalledTimes(1);
    const rangeArg = spy.mock.calls[0]?.[1] as any;
    expect(rangeArg?.start?.row).toBe(0);
    expect(rangeArg?.start?.col).toBe(0);
    expect(rangeArg?.end?.row).toBe(0);
    expect(rangeArg?.end?.col).toBe(DEFAULT_GRID_LIMITS.maxCols - 1);
  });

  it("initializes dialog controls from snake_case (formula-model) formatting keys", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", {
      alignment: { wrap_text: true },
      font: { size_100pt: 1200 },
      number_format: "0%",
    });

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [],
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    const wrap = dialog!.querySelector<HTMLInputElement>('[data-testid="format-cells-wrap"]');
    expect(wrap?.checked).toBe(true);

    const number = dialog!.querySelector<HTMLSelectElement>('[data-testid="format-cells-number"]');
    expect(number?.value).toBe("percent");

    const fontSize = dialog!.querySelector<HTMLInputElement>('[data-testid="format-cells-font-size"]');
    expect(fontSize?.value).toBe("12");
  });

  it("applies a custom number format string when __custom__ is selected", () => {
    const doc = new DocumentController();

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [],
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    const number = dialog!.querySelector<HTMLSelectElement>('[data-testid="format-cells-number"]');
    expect(number).not.toBeNull();
    number!.value = "__custom__";
    number!.dispatchEvent(new Event("change", { bubbles: true }));

    const customInput = dialog!.querySelector<HTMLInputElement>('[data-testid="format-cells-number-custom"]');
    expect(customInput).not.toBeNull();
    customInput!.value = "#,##0.00";

    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(doc.getCellFormat("Sheet1", "A1").numberFormat).toBe("#,##0.00");
  });

  it("initializes Number->Custom UI from the active cell's custom number format", () => {
    const doc = new DocumentController();
    doc.setRangeFormat("Sheet1", "A1", { numberFormat: "#,##0.00" });

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [],
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    const number = dialog!.querySelector<HTMLSelectElement>('[data-testid="format-cells-number"]');
    expect(number?.value).toBe("__custom__");

    const customInput = dialog!.querySelector<HTMLInputElement>('[data-testid="format-cells-number-custom"]');
    expect(customInput?.value).toBe("#,##0.00");

    const customRow = customInput?.closest<HTMLElement>(".format-cells-dialog__row");
    expect(customRow?.style.display).not.toBe("none");
    // The dialog should focus the code input when opening with an existing custom format.
    expect(document.activeElement).toBe(customInput);
  });

  it("preserves whitespace in custom number format codes", () => {
    const doc = new DocumentController();

    openFormatCellsDialog({
      isEditing: () => false,
      getDocument: () => doc,
      getSheetId: () => "Sheet1",
      getActiveCell: () => ({ row: 0, col: 0 }),
      getSelectionRanges: () => [],
      focusGrid: () => {},
    });

    const dialog = document.querySelector<HTMLDialogElement>("dialog.format-cells-dialog");
    expect(dialog).not.toBeNull();

    const number = dialog!.querySelector<HTMLSelectElement>('[data-testid="format-cells-number"]');
    number!.value = "__custom__";
    number!.dispatchEvent(new Event("change", { bubbles: true }));

    const customInput = dialog!.querySelector<HTMLInputElement>('[data-testid="format-cells-number-custom"]');
    customInput!.value = "0.00 ";

    dialog!.querySelector<HTMLButtonElement>('[data-testid="format-cells-apply"]')!.click();

    expect(doc.getCellFormat("Sheet1", "A1").numberFormat).toBe("0.00 ");
  });
});
