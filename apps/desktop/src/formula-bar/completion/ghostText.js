/**
 * UI helper: compute "ghost text" (inline suggestion) for the formula bar.
 *
 * The completion engine returns full suggestion text; the UI typically wants to
 * show only the "tail" after the cursor (like VS Code inline suggestions).
 */

/**
 * @param {string} currentInput
 * @param {number} cursorPosition
 * @param {{ text: string } | null | undefined} suggestion
 * @returns {string}
 */
export function getGhostText(currentInput, cursorPosition, suggestion) {
  const suggestionText = suggestion?.text;
  if (typeof suggestionText !== "string" || suggestionText.length === 0) return "";

  const cursor = clampCursor(currentInput, cursorPosition);
  const prefix = currentInput.slice(0, cursor);
  const suffix = currentInput.slice(cursor);

  // Only support "pure insertion at cursor" for now.
  if (!suggestionText.startsWith(prefix)) return "";
  if (suffix && !suggestionText.endsWith(suffix)) return "";

  return suggestionText.slice(cursor, suggestionText.length - suffix.length);
}

/**
 * Apply a suggestion as if the user pressed Tab.
 *
 * @param {string} currentInput
 * @param {number} cursorPosition
 * @param {{ text: string } | null | undefined} suggestion
 * @returns {{ text: string, cursorPosition: number }}
 */
export function acceptSuggestion(currentInput, cursorPosition, suggestion) {
  const suggestionText = suggestion?.text;
  if (typeof suggestionText !== "string" || suggestionText.length === 0) {
    return { text: currentInput, cursorPosition: clampCursor(currentInput, cursorPosition) };
  }

  const cursor = clampCursor(currentInput, cursorPosition);
  const ghost = getGhostText(currentInput, cursor, { text: suggestionText });
  const newCursor = ghost ? cursor + ghost.length : suggestionText.length;
  return { text: suggestionText, cursorPosition: newCursor };
}

function clampCursor(input, cursorPosition) {
  if (!Number.isInteger(cursorPosition)) return input.length;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > input.length) return input.length;
  return cursorPosition;
}
