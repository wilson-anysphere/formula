import type { RichText } from "./types.js";

/**
 * MVP editing preservation rule:
 * - In-cell editing operates on plain text.
 * - If the user commits without changing the text, keep the original rich runs.
 * - If the user changes the text, convert to plain rich text (empty runs).
 */
export function applyPlainTextEdit(
  original: RichText | null | undefined,
  editedText: string,
): RichText;

