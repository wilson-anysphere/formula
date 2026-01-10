/**
 * @typedef {import("./types.js").PasteSpecialMode} PasteSpecialMode
 * @typedef {{ mode: PasteSpecialMode, label: string }} PasteSpecialMenuItem
 */

import { t } from "../i18n/index.js";

/**
 * UI skeleton for Paste Special.
 *
 * The spreadsheet UI can render these items in a context menu / command palette
 * and pass the selected `mode` into the paste application logic.
 */
/**
 * @returns {PasteSpecialMenuItem[]}
 */
export function getPasteSpecialMenuItems() {
  return [
    { mode: "all", label: t("clipboard.pasteSpecial.paste") },
    { mode: "values", label: t("clipboard.pasteSpecial.pasteValues") },
    { mode: "formulas", label: t("clipboard.pasteSpecial.pasteFormulas") },
    { mode: "formats", label: t("clipboard.pasteSpecial.pasteFormats") },
  ];
}
