/**
 * Find the end position (exclusive) for a bracketed segment starting at `startIndex`.
 *
 * Excel uses `]]` to encode a literal `]` inside structured references and external workbook
 * prefixes, which is ambiguous with nested closure (e.g. `[[Col]]`).
 *
 * This matcher prefers treating `]]` as an escaped `]` but will backtrack if that interpretation
 * makes it impossible to close all brackets before `limit`.
 *
 * This is intentionally a small, dependency-free helper that can be reused by the lightweight
 * formula partial parser and completion utilities without pulling in the full formula tokenizer.
 *
 * @param {string} src
 * @param {number} startIndex
 * @param {number} limit exclusive upper bound for scanning
 * @returns {number | null}
 */
export function findMatchingBracketEnd(src, startIndex, limit) {
  if (typeof src !== "string") return null;
  const max = typeof limit === "number" && Number.isFinite(limit) ? Math.max(0, Math.min(src.length, limit)) : src.length;
  if (startIndex < 0 || startIndex >= max) return null;
  if (src[startIndex] !== "[") return null;

  let i = startIndex;
  let depth = 0;
  /** @type {Array<{ i: number, depth: number }>} */
  const escapeChoices = [];

  const backtrack = () => {
    const choice = escapeChoices.pop();
    if (!choice) return false;
    i = choice.i;
    depth = choice.depth;
    // Reinterpret the first `]` of the `]]` pair as a real closing bracket.
    depth -= 1;
    i += 1;
    return true;
  };

  while (true) {
    if (i >= max) {
      if (!backtrack()) return null;
      continue;
    }

    const ch = src[i];
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }
    if (ch === "]") {
      if (src[i + 1] === "]" && depth > 0 && i + 1 < max) {
        // Prefer treating `]]` as an escaped literal `]`. Record a choice point so we can
        // reinterpret it as a real closing bracket if we later fail to close everything.
        escapeChoices.push({ i, depth });
        i += 2;
        continue;
      }
      depth -= 1;
      i += 1;
      if (depth === 0) return i;
      if (depth < 0) {
        // Too many closing brackets; try reinterpreting an earlier escape.
        if (!backtrack()) return null;
      }
      continue;
    }

    i += 1;
  }
}

