/**
 * Fast, model-agnostic token estimate. Works well enough for enforcing a budget
 * in a UI context manager; exact tokenization is model-specific and can be added
 * later behind the same interface.
 *
 * @param {string} text
 */
export function estimateTokens(text) {
  if (!text) return 0;
  // A common approximation: 4 chars/token for English-like text.
  return Math.ceil(text.length / 4);
}

/**
 * Trim a string to a maximum estimated token count.
 * @param {string} text
 * @param {number} maxTokens
 */
export function trimToTokenBudget(text, maxTokens) {
  if (maxTokens <= 0) return "";
  const estimate = estimateTokens(text);
  if (estimate <= maxTokens) return text;

  const maxChars = Math.max(maxTokens * 4, 0);
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars) + "\n…(trimmed to fit token budget)…";
}

/**
 * @param {{ key: string, text: string, priority: number }[]} sections
 * @param {number} maxTokens
 */
export function packSectionsToTokenBudget(sections, maxTokens) {
  const ordered = sections.slice().sort((a, b) => b.priority - a.priority);
  let remaining = maxTokens;
  /** @type {{ key: string, text: string }[]} */
  const packed = [];

  for (const section of ordered) {
    if (remaining <= 0) break;
    const trimmed = trimToTokenBudget(section.text, remaining);
    const used = estimateTokens(trimmed);
    remaining -= used;
    packed.push({ key: section.key, text: trimmed });
  }

  return packed;
}
