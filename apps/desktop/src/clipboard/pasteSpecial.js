/**
 * @typedef {import("./types.js").PasteSpecialMode} PasteSpecialMode
 * @typedef {{ mode: PasteSpecialMode, label: string }} PasteSpecialMenuItem
 */

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
    { mode: "all", label: "Paste" },
    { mode: "values", label: "Paste Values" },
    { mode: "formulas", label: "Paste Formulas" },
    { mode: "formats", label: "Paste Formats" },
  ];
}
