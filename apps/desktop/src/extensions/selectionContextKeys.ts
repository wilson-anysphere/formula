import { cellToA1 } from "../selection/a1";
import type { SelectionState, SelectionType } from "../selection/types";

export type SelectionContextKeys = {
  selectionType: SelectionType;
  /**
   * `true` when the selection is anything other than a single cell.
   *
   * This aligns with menu/keybinding semantics (row/column/all selections should count),
   * and avoids "multi selection of single cells" being treated as no selection.
   */
  hasSelection: boolean;
  isSingleCell: boolean;
  isMultiRange: boolean;
  activeCellA1: string;
};

export function deriveSelectionContextKeys(selection: SelectionState): SelectionContextKeys {
  const selectionType = selection.type;
  const isSingleCell = selectionType === "cell";
  return {
    selectionType,
    hasSelection: !isSingleCell,
    isSingleCell,
    isMultiRange: selectionType === "multi",
    activeCellA1: cellToA1(selection.active),
  };
}

