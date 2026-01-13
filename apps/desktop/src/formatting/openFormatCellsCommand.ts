import type { FormatCellsDialogHost } from "./openFormatCellsDialog.js";
import { openFormatCellsDialog } from "./openFormatCellsDialog.js";

/**
 * Builds a zero-arg callback that opens the Format Cells dialog for the provided host.
 *
 * This is used by the `format.openFormatCells` command (Ctrl/Cmd+1) and can be shared
 * by other UI entrypoints that need to open the dialog with the same host wiring.
 */
export function createOpenFormatCells(host: FormatCellsDialogHost): () => void {
  return () => openFormatCellsDialog(host);
}
