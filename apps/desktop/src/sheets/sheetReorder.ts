import type { SheetVisibility } from "./workbookSheetStore";

/**
 * Minimal sheet shape required for computing workbook order changes.
 *
 * We intentionally depend only on `{ id, visibility }` so the reorder logic can
 * be shared between the React sheet tab strip and any other sheet metadata
 * stores/models.
 */
export type SheetOrderEntry = {
  id: string;
  visibility: SheetVisibility;
};

export type SheetTabDropTarget =
  | {
      /**
       * Drop the dragged sheet so it appears *before* the target sheet in the
       * underlying workbook order (visible + hidden).
       */
      kind: "before";
      targetSheetId: string;
    }
  | {
      /**
       * Drop the dragged sheet at the end of the visible tab strip.
       *
       * This maps to "after the last visible sheet" in the full workbook order,
       * but still *before* any trailing hidden sheets (Excel semantics).
       */
      kind: "end";
    };

/**
 * Compute the absolute destination index in the full workbook sheet order when
 * reordering via the visible sheet tab strip.
 *
 * The tab strip only renders `visibility === "visible"` sheets, but the
 * underlying workbook order includes hidden sheets. To match Excel semantics we
 *:
 *   1) Remove the dragged sheet from the full order
 *   2) Insert it at the destination position in the full order
 *
 * This function returns the `newIndex` suitable for passing to a "move sheet"
 * primitive that performs a remove+insert reorder on the full sheet list.
 */
export function computeWorkbookSheetMoveIndex(params: {
  sheets: readonly SheetOrderEntry[];
  fromSheetId: string;
  dropTarget: SheetTabDropTarget;
}): number | null {
  const { sheets, fromSheetId, dropTarget } = params;
  if (sheets.length === 0) return null;

  const fromIndex = sheets.findIndex((s) => s.id === fromSheetId);
  if (fromIndex < 0) return null;

  if (dropTarget.kind === "end") {
    // Place the dragged sheet at the end of the visible strip, which corresponds
    // to the index of the last visible sheet in the full workbook order. This
    // keeps any trailing hidden sheets at the end of the workbook order.
    let lastVisibleIndex = -1;
    for (let i = sheets.length - 1; i >= 0; i -= 1) {
      if (sheets[i]?.visibility === "visible") {
        lastVisibleIndex = i;
        break;
      }
    }
    if (lastVisibleIndex < 0) {
      // Defensive fallback: if no visible sheets exist, treat this as a move to
      // the end of the sheet list.
      return sheets.length - 1;
    }
    return lastVisibleIndex;
  }

  const targetIndex = sheets.findIndex((s) => s.id === dropTarget.targetSheetId);
  if (targetIndex < 0) return null;

  // The tab strip drop semantics are "insert before target". If the dragged
  // sheet currently appears before the target, removing it shifts the target
  // left by one.
  const toIndex = fromIndex < targetIndex ? targetIndex - 1 : targetIndex;
  return Math.max(0, Math.min(toIndex, sheets.length - 1));
}
