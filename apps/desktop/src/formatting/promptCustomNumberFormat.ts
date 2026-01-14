import type { DocumentController } from "../document/documentController.js";
import type { CellRange } from "./toolbar.js";

export type ApplyFormattingToSelection = (
  label: string,
  fn: (doc: DocumentController, sheetId: string, ranges: CellRange[]) => void | boolean,
) => void;

export type ShowInputBox = (options: {
  prompt?: string;
  value?: string;
  placeHolder?: string;
}) => Promise<string | null>;

export async function promptAndApplyCustomNumberFormat(options: {
  isEditing: () => boolean;
  showInputBox: ShowInputBox;
  getActiveCellNumberFormat: () => string | null;
  applyFormattingToSelection: ApplyFormattingToSelection;
}): Promise<void> {
  if (options.isEditing()) return;

  const seed = options.getActiveCellNumberFormat() ?? "";
  const input = await options.showInputBox({
    prompt: "Custom number format code",
    value: seed,
    placeHolder: "General",
  });
  if (input == null) return;

  // Avoid applying formatting if the user started editing while the prompt was open.
  if (options.isEditing()) return;

  // Preserve the exact user-entered format code. (Excel number formats can contain spaces,
  // so avoid trimming beyond what we need for "empty"/"General" detection.)
  const trimmed = input.trim();
  const desired = !trimmed || trimmed.toLowerCase() === "general" ? null : input;

  options.applyFormattingToSelection("Number format", (doc, sheetId, ranges) => {
    let applied = true;
    for (const range of ranges) {
      const ok = doc.setRangeFormat(sheetId, range, { numberFormat: desired }, { label: "Number format" });
      if (ok === false) applied = false;
    }
    return applied;
  });
}
