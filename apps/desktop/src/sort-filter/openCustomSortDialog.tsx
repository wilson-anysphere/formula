import React from "react";
import { createRoot } from "react-dom/client";

import type { DocumentController } from "../document/documentController.js";
import { showToast } from "../extensions/ui.js";
import { markKeybindingBarrier } from "../keybindingBarrier.js";
import type { Range } from "../selection/types";
import type { SpreadsheetValue } from "../spreadsheet/evaluateFormula";

import { SortDialog } from "./SortDialog";
import * as sortSelection from "./sortSelection";
import type { SortSpec } from "./types";

export type CustomSortDialogHost = {
  isEditing: () => boolean;
  getDocument: () => DocumentController;
  getSheetId: () => string;
  getSelectionRanges: () => Range[];
  /**
   * Returns a computed value for sorting/headers. This should match what the grid shows
   * (e.g. formulas evaluated to their computed value).
   */
  getCellValue: (sheetId: string, cell: { row: number; col: number }) => SpreadsheetValue;
  focusGrid: () => void;
};

function showModal(dialog: HTMLDialogElement): void {
  // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
  if (typeof dialog.showModal === "function") {
    // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
    dialog.showModal();
    return;
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
}): { columns: Array<{ index: number; name: string }>; headerDetected: boolean } {
  const selection = normalizeRange(params.selection);
  const width = selection.endCol - selection.startCol + 1;
  const labels = Array.from({ length: width }, (_, idx) => colIndexToLabel(idx));

  const rawHeaderTexts = labels.map((_fallback, idx) => {
    const value = params.host.getCellValue(params.sheetId, { row: selection.startRow, col: selection.startCol + idx });
    const text = value == null ? "" : String(value);
    return text.trim();
  });

  const headerDetected = rawHeaderTexts.some((t) => t !== "");
  const headerTexts = rawHeaderTexts.map((text, idx) => (text ? text : labels[idx]!));

  return {
    columns: labels.map((fallback, idx) => ({
      index: idx,
      name: headerDetected ? headerTexts[idx]! : fallback,
    })),
    headerDetected,
  };
}

export function openCustomSortDialog(host: CustomSortDialogHost): void {
  if (host.isEditing()) return;

  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) return;

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
  const { columns, headerDetected } = deriveColumnsFromSelection({ host, sheetId, selection });

  const initial: SortSpec = {
    keys: [{ column: 0, order: "ascending" }],
    hasHeader: headerDetected,
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

  root.render(
    <SortDialog
      columns={columns}
      initial={initial}
      onCancel={() => closeDialog(dialog, "cancel")}
      onApply={(spec) => {
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
