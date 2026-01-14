import type { DocumentController } from "../document/documentController.js";
import { t } from "../i18n/index.js";
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

  const numberFormatLabel = t("quickPick.numberFormat.placeholder");
  const seed = options.getActiveCellNumberFormat() ?? "";
  const input = await options.showInputBox({
    prompt: t("prompt.customNumberFormat.code"),
    value: seed,
    placeHolder: t("command.format.numberFormat.general"),
  });
  if (input == null) return;

  // Avoid applying formatting if the user started editing while the prompt was open.
  if (options.isEditing()) return;

  // Preserve the exact user-entered format code. (Excel number formats can contain spaces,
  // so avoid trimming beyond what we need for "empty"/"General" detection.)
  const trimmed = input.trim();
  const normalized = trimmed.toLowerCase();
  const localizedGeneral = t("command.format.numberFormat.general").trim().toLowerCase();
  const desired =
    !trimmed || normalized === "general" || (localizedGeneral && normalized === localizedGeneral) ? null : input;

  options.applyFormattingToSelection(numberFormatLabel, (doc, sheetId, ranges) => {
    let applied = true;
    for (const range of ranges) {
      const ok = doc.setRangeFormat(sheetId, range, { numberFormat: desired }, { label: numberFormatLabel });
      if (ok === false) applied = false;
    }
    return applied;
  });
}
