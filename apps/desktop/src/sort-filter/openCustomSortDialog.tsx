import React from "react";
import { createRoot } from "react-dom/client";

import type { DocumentController } from "../document/documentController.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";
import { showToast } from "../extensions/ui.js";
import { markKeybindingBarrier } from "../keybindingBarrier.js";
import type { Range } from "../selection/types";
import type { SpreadsheetValue } from "../spreadsheet/evaluateFormula";

import { SortDialog } from "./SortDialog";
import * as sortSelection from "./sortSelection";
import type { SortSpec } from "./types";

export type CustomSortDialogHost = {
  isEditing: () => boolean;
  /**
   * Optional read-only indicator (used in collab viewer/commenter sessions).
   *
   * Sorting mutates cell contents, so it must remain blocked in read-only roles even though
   * other view-local operations (like AutoFilter row hiding) may still be allowed.
   */
  isReadOnly?: () => boolean;
  getDocument: () => DocumentController;
  getSheetId: () => string;
  getSelectionRanges: () => Range[];
  /**
   * Returns a computed value for sorting/headers. This should match what the grid shows
   * (e.g. formulas evaluated to their computed value).
   */
  getCellValue: (sheetId: string, cell: { row: number; col: number }) => SpreadsheetValue;
  /**
   * Optional hook to infer why a cell edit is blocked (permission vs missing encryption key).
   *
   * This is provided by SpreadsheetApp so sort operations can show encryption-aware rejection
   * toasts instead of silently no-op'ing.
   */
  inferCollabEditRejection?: (cell: {
    sheetId: string;
    row: number;
    col: number;
  }) => { rejectionReason: "permission" | "encryption" | "unknown"; encryptionKeyId?: string; encryptionPayloadUnsupported?: boolean };
  focusGrid: () => void;
};

function showModal(dialog: HTMLDialogElement): void {
  // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
  if (typeof dialog.showModal === "function") {
    try {
      // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
      dialog.showModal();
      return;
    } catch {
      // Fall through to non-modal open attribute.
    }
  }
  // jsdom fallback: `open` attribute is enough for our tests.
  dialog.setAttribute("open", "true");
}

function isDialogOpen(dialog: HTMLDialogElement): boolean {
  // @ts-expect-error - jsdom typing mismatch.
  return dialog.open === true || dialog.hasAttribute("open");
}

function closeDialog(dialog: HTMLDialogElement, returnValue: string): void {
  if (!isDialogOpen(dialog)) return;
  // @ts-expect-error - HTMLDialogElement.close() not implemented in jsdom.
  if (typeof dialog.close === "function") {
    // @ts-expect-error - HTMLDialogElement.close() not implemented in jsdom.
    dialog.close(returnValue);
    return;
  }
  // jsdom fallback: emulate the returnValue contract + close event.
  // @ts-expect-error - returnValue not modeled on jsdom's dialog typings.
  dialog.returnValue = returnValue;
  dialog.removeAttribute("open");
  dialog.dispatchEvent(new Event("close"));
}

function trapTabNavigation(dialog: HTMLDialogElement, event: KeyboardEvent): void {
  if (event.key !== "Tab") return;
  const focusables = Array.from(
    dialog.querySelectorAll<HTMLElement>(
      'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
    ),
  ).filter((el) => el.getAttribute("aria-hidden") !== "true");
  if (focusables.length === 0) return;
  const first = focusables[0]!;
  const last = focusables[focusables.length - 1]!;
  const active = document.activeElement as HTMLElement | null;
  if (!active) return;

  if (event.shiftKey) {
    if (active === first) {
      event.preventDefault();
      last.focus();
    }
    return;
  }

  if (active === last) {
    event.preventDefault();
    first.focus();
  }
}

function normalizeRange(range: Range): Range {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

function colIndexToLabel(index: number): string {
  // A1-style column naming.
  let n = Math.max(0, Math.trunc(index)) + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out || "A";
}

function deriveColumnsFromSelection(params: {
  host: CustomSortDialogHost;
  sheetId: string;
  selection: Range;
}): { headerColumns: Array<{ index: number; name: string }>; fallbackColumns: Array<{ index: number; name: string }>; initialHasHeader: boolean } {
  const selection = normalizeRange(params.selection);
  const width = selection.endCol - selection.startCol + 1;
  // Use the sheet's column letters for the selection (e.g. D/E/F when selecting columns D–F).
  const fallbackLabels = Array.from({ length: width }, (_, idx) => colIndexToLabel(selection.startCol + idx));
  const fallbackColumns = fallbackLabels.map((name, index) => ({ index, name }));

  const rawHeaderTexts = fallbackLabels.map((_fallback, idx) => {
    const value = params.host.getCellValue(params.sheetId, {
      row: selection.startRow,
      col: selection.startCol + idx,
    });
    const text = value == null ? "" : String(value);
    return text.trim();
  });

  const headerTexts = rawHeaderTexts.map((text, idx) => (text ? text : fallbackLabels[idx]!));
  const headerColumns = fallbackLabels.map((_fallback, idx) => ({
    index: idx,
    name: headerTexts[idx]!,
  }));

  // Heuristic: if the first row contains any non-empty values, assume it's a header row.
  // Users can always toggle this off if the selection does not include headers.
  const initialHasHeader = rawHeaderTexts.some((text) => text !== "");

  return { headerColumns, fallbackColumns, initialHasHeader };
}

export function openCustomSortDialog(host: CustomSortDialogHost): void {
  if (host.isEditing()) return;

  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) return;

  if (host.isReadOnly?.() === true) {
    showCollabEditRejectedToast([{ rejectionKind: "sort", rejectionReason: "permission" }]);
    try {
      host.focusGrid();
    } catch {
      // ignore
    }
    return;
  }

  const selectionRanges = host.getSelectionRanges();
  if (selectionRanges.length !== 1) {
    showToast("Select a single range to sort.", "warning");
    return;
  }

  const selection = normalizeRange(selectionRanges[0]!);
  const width = selection.endCol - selection.startCol + 1;
  const height = selection.endRow - selection.startRow + 1;
  const area = width * height;
  if (area > sortSelection.DEFAULT_SORT_CELL_LIMIT) {
    showToast("Selection is too large to sort. Try selecting fewer cells.", "warning");
    return;
  }

  const sheetId = host.getSheetId();
  const { headerColumns, fallbackColumns, initialHasHeader } = deriveColumnsFromSelection({ host, sheetId, selection });

  const initial: SortSpec = {
    keys: [{ column: 0, order: "ascending" }],
    hasHeader: initialHasHeader,
  };

  const dialog = document.createElement("dialog");
  dialog.className = "dialog custom-sort-dialog";
  dialog.dataset.testid = "custom-sort-dialog";
  dialog.setAttribute("aria-label", "Custom Sort");
  markKeybindingBarrier(dialog);

  const content = document.createElement("div");
  dialog.appendChild(content);

  document.body.appendChild(dialog);

  const root = createRoot(content);

  const cleanup = () => {
    try {
      root.unmount();
    } catch {
      // Best-effort.
    }
    dialog.remove();
    host.focusGrid();
  };

  dialog.addEventListener(
    "close",
    () => {
      cleanup();
    },
    { once: true },
  );

  dialog.addEventListener("cancel", (e) => {
    e.preventDefault();
    closeDialog(dialog, "cancel");
  });

  // Trap Tab navigation within the modal so focus doesn't escape back to the grid/ribbon.
  dialog.addEventListener("keydown", (event) => trapTabNavigation(dialog, event));

  root.render(
    <SortDialog
      columns={headerColumns}
      fallbackColumns={fallbackColumns}
      initial={initial}
      onCancel={() => closeDialog(dialog, "cancel")}
      onApply={(spec) => {
        // Sorting writes the full selection range back into the document. In collab/protected/encrypted
        // contexts, `DocumentController` filters writes per-cell via `canEditCell`. For sort this is
        // unsafe because it can corrupt row integrity if any cell is blocked. Preflight the selection
        // and abort with a toast instead of attempting a partial sort.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const canEditCell = (host.getDocument() as any)?.canEditCell as
          | ((cell: { sheetId: string; row: number; col: number }) => boolean)
          | null
          | undefined;
        if (typeof canEditCell === "function") {
          let blocked: { row: number; col: number } | null = null;
          try {
            outer: for (let row = selection.startRow; row <= selection.endRow; row += 1) {
              for (let col = selection.startCol; col <= selection.endCol; col += 1) {
                if (!canEditCell.call(host.getDocument(), { sheetId, row, col })) {
                  blocked = { row, col };
                  break outer;
                }
              }
            }
          } catch {
            blocked = null;
          }
          if (blocked) {
            const rejection = (() => {
              if (typeof host.inferCollabEditRejection === "function") {
                const inferred = host.inferCollabEditRejection({ sheetId, row: blocked.row, col: blocked.col });
                if (inferred && typeof inferred === "object" && typeof (inferred as any).rejectionReason === "string") {
                  return inferred;
                }
              }
              return { rejectionReason: "permission" as const };
            })();
            showCollabEditRejectedToast([
              {
                rejectionKind: "sort",
                ...rejection,
              },
            ]);
            return;
          }
        }

        const ok = sortSelection.applySortSpecToSelection({
          doc: host.getDocument(),
          sheetId,
          selection,
          spec,
          getCellValue: (cell) => host.getCellValue(sheetId, cell),
          maxCells: sortSelection.DEFAULT_SORT_CELL_LIMIT,
          label: "Sort",
        });
        if (!ok) {
          showToast("Could not sort selection. Try selecting fewer cells.", "warning");
          return;
        }
        closeDialog(dialog, "ok");
      }}
    />,
  );

  showModal(dialog);

  // Focus the first control rendered by SortDialog (checkbox).
  const schedule =
    typeof requestAnimationFrame === "function"
      ? requestAnimationFrame
      : (cb: FrameRequestCallback) => window.setTimeout(() => cb(Date.now()), 0);
  schedule(() => {
    const el = dialog.querySelector<HTMLElement>("input, select, button");
    try {
      el?.focus();
    } catch {
      // Best-effort focus.
    }
  });
}

export function handleCustomSortCommand(commandId: string, host: CustomSortDialogHost): boolean {
  if (commandId !== "home.editing.sortFilter.customSort" && commandId !== "data.sortFilter.sort.customSort") return false;
  openCustomSortDialog(host);
  return true;
}
