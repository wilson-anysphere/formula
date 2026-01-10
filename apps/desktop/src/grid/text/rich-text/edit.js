/**
 * MVP editing preservation rule:
 * - In-cell editing operates on plain text.
 * - If the user commits without changing the text, keep the original rich runs.
 * - If the user changes the text, convert to plain (single-run) rich text.
 *
 * @param {import('./types.js').RichText | null | undefined} original
 * @param {string} editedText
 * @returns {import('./types.js').RichText}
 */
export function applyPlainTextEdit(original, editedText) {
  if (original && original.text === editedText) {
    return original;
  }
  return { text: editedText, runs: [] };
}

