import type { DocumentController } from "../document/documentController.js";
import { t } from "../i18n/index.js";
import type { CellRange } from "./toolbar.js";
import { isValidExcelNumberFormatCode } from "./numberFormat.js";

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
  /**
   * Seed value for the input prompt. Callers should return the current selection's effective
   * number format when it is a single consistent code; otherwise return null to leave the
   * prompt empty.
   */
  getSelectionNumberFormat: () => string | null;
  applyFormattingToSelection: ApplyFormattingToSelection;
  showToast?: (message: string, type?: "info" | "warning" | "error") => void;
}): Promise<void> {
  if (options.isEditing()) return;

  const numberFormatLabel = t("quickPick.numberFormat.placeholder");
  const seed = options.getSelectionNumberFormat() ?? "";
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

  if (typeof desired === "string") {
    // Best-effort validation: catch obvious syntax errors before applying.
    if (!isValidExcelNumberFormatCode(desired)) {
      options.showToast?.(t("toast.customNumberFormat.invalid"), "warning");
      return;
    }
  }

  options.applyFormattingToSelection(numberFormatLabel, (doc, sheetId, ranges) => {
    let applied = true;
    for (const range of ranges) {
      const ok = doc.setRangeFormat(sheetId, range, { numberFormat: desired }, { label: numberFormatLabel });
      if (ok === false) applied = false;
    }
    return applied;
  });
}
