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
});
