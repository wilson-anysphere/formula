import { normalizeFormulaText } from "./normalize.js";
import { tokenizeFormula } from "./tokenize.js";

/**
 * @typedef {import("./tokenize.js").Token} Token
 * @typedef {"equal" | "insert" | "delete"} DiffOpType
 * @typedef {{ type: DiffOpType, tokens: Token[] }} DiffOp
 */

/**
 * @param {Token} a
 * @param {Token} b
 * @param {{ normalize: boolean }} opts
 */
function tokenEquals(a, b, opts) {
  if (a.type !== b.type) return false;
  // Excel formulas are generally case-insensitive for identifiers/cell refs.
  // Treat case-only edits as cosmetic when normalization is enabled.
  if (opts.normalize && a.type === "ident") {
    return a.value.toUpperCase() === b.value.toUpperCase();
  }
  return a.value === b.value;
}

/**
 * Tokenize a formula for diffing.
 *
 * - Removes the trailing EOF token produced by `tokenizeFormula`.
 * - Returns an empty array for `null` / empty formulas.
 *
 * @param {string | null} formula
 * @returns {Token[]}
 */
function tokenizeForDiff(formula) {
  if (formula == null) return [];
  try {
    const tokens = tokenizeFormula(formula);
    // The EOF token is useful for parsing but noisy for UI-level diffs.
    if (tokens.length > 0 && tokens[tokens.length - 1]?.type === "eof") tokens.pop();
    return tokens;
  } catch {
    // Tokenization can fail on incomplete formulas (e.g. unterminated string
    // literals). For diffs/history UI we prefer a best-effort token stream
    // rather than throwing.
    const text = String(formula);
    if (text.startsWith("=")) {
      const rest = text.slice(1);
      return rest ? [{ type: "op", value: "=" }, { type: "ident", value: rest }] : [{ type: "op", value: "=" }];
    }
    return [{ type: "ident", value: text }];
  }
}

/**
 * @template T
 * @param {T[]} a
 * @param {T[]} b
 * @param {(x: T, y: T) => boolean} equals
 * @returns {Array<{ type: DiffOpType, token: T }>}
 */
function myersDiff(a, b, equals) {
  const n = a.length;
  const m = b.length;

  if (n === 0 && m === 0) return [];
  if (n === 0) return b.map((token) => ({ type: "insert", token }));
  if (m === 0) return a.map((token) => ({ type: "delete", token }));

  const max = n + m;
  const offset = max;
  /** @type {number[]} */
  let v = new Array(2 * max + 1).fill(0);
  /** @type {number[][]} */
  const trace = [];

  for (let d = 0; d <= max; d += 1) {
    trace.push(v.slice());

    for (let k = -d; k <= d; k += 2) {
      const kIndex = offset + k;
      let x;

      if (k === -d || (k !== d && v[kIndex - 1] < v[kIndex + 1])) {
        // Down: insert into A (advance in B).
        x = v[kIndex + 1];
      } else {
        // Right: delete from A.
        x = v[kIndex - 1] + 1;
      }

      let y = x - k;
      while (x < n && y < m && equals(a[x], b[y])) {
        x += 1;
        y += 1;
      }
      v[kIndex] = x;

      if (x >= n && y >= m) {
        return backtrackMyers(trace, a, b, offset);
      }
    }
  }

  // Unreachable, but keep a safe fallback.
  /** @type {Array<{ type: DiffOpType, token: T }>} */
  const fallback = [];
  for (const token of a) fallback.push({ type: "delete", token });
  for (const token of b) fallback.push({ type: "insert", token });
  return fallback;

  /**
   * @param {number[][]} trace
   * @param {T[]} a
   * @param {T[]} b
   * @param {number} offset
   */
  function backtrackMyers(trace, a, b, offset) {
    let x = a.length;
    let y = b.length;
    /** @type {Array<{ type: DiffOpType, token: T }>} */
    const edits = [];

    for (let d = trace.length - 1; d > 0; d -= 1) {
      const v = trace[d];
      const k = x - y;
      const kIndex = offset + k;

      let prevK;
      if (k === -d || (k !== d && v[kIndex - 1] < v[kIndex + 1])) {
        prevK = k + 1;
      } else {
        prevK = k - 1;
      }

      const prevX = v[offset + prevK];
      const prevY = prevX - prevK;

      while (x > prevX && y > prevY) {
        edits.push({ type: "equal", token: a[x - 1] });
        x -= 1;
        y -= 1;
      }

      if (x === prevX) {
        edits.push({ type: "insert", token: b[prevY] });
        y -= 1;
      } else {
        edits.push({ type: "delete", token: a[prevX] });
        x -= 1;
      }
    }

    // d === 0: the remaining path is only diagonal moves (common prefix).
    while (x > 0 && y > 0) {
      edits.push({ type: "equal", token: a[x - 1] });
      x -= 1;
      y -= 1;
    }
    while (x > 0) {
      edits.push({ type: "delete", token: a[x - 1] });
      x -= 1;
    }
    while (y > 0) {
      edits.push({ type: "insert", token: b[y - 1] });
      y -= 1;
    }

    edits.reverse();
    return edits;
  }
}

/**
 * @param {Array<{ type: DiffOpType, token: Token }>} edits
 * @returns {DiffOp[]}
 */
function coalesceEdits(edits) {
  /** @type {DiffOp[]} */
  const ops = [];

  for (const edit of edits) {
    const last = ops[ops.length - 1];
    if (!last || last.type !== edit.type) {
      ops.push({ type: edit.type, tokens: [edit.token] });
      continue;
    }
    last.tokens.push(edit.token);
  }

  return ops;
}

/**
 * @template T
 * @param {T[]} a
 * @param {T[]} b
 * @param {(x: T, y: T) => boolean} equals
 */
function commonPrefixLength(a, b, equals) {
  const len = Math.min(a.length, b.length);
  let i = 0;
  while (i < len && equals(a[i], b[i])) i += 1;
  return i;
}

/**
 * @template T
 * @param {T[]} a
 * @param {T[]} b
 * @param {(x: T, y: T) => boolean} equals
 * @param {number} prefixLen
 */
function commonSuffixLength(a, b, equals, prefixLen) {
  const max = Math.min(a.length, b.length) - prefixLen;
  let i = 0;
  while (i < max && equals(a[a.length - 1 - i], b[b.length - 1 - i])) i += 1;
  return i;
}

/**
 * Formula-aware diff returning token operations so callers can render changes with
 * syntax highlighting.
 *
 * By default (`opts.normalize !== false`), formulas are lightly normalized before
 * tokenization:
 * - trimmed
 * - ensured to have a leading `=`
 *
 * @param {string | null} oldFormula
 * @param {string | null} newFormula
 * @param {{ normalize?: boolean } | undefined} opts
 * @returns {{ equal: boolean, ops: DiffOp[] }}
 */
export function diffFormula(oldFormula, newFormula, opts) {
  const normalize = opts?.normalize !== false;

  const oldText = normalize ? normalizeFormulaText(oldFormula) : oldFormula?.trim() || null;
  const newText = normalize ? normalizeFormulaText(newFormula) : newFormula?.trim() || null;

  const oldTokens = tokenizeForDiff(oldText);
  const newTokens = tokenizeForDiff(newText);

  const equals = (a, b) => tokenEquals(a, b, { normalize });

  const equal =
    oldTokens.length === newTokens.length &&
    oldTokens.every((t, i) => equals(t, newTokens[i]));

  // Trim common prefix/suffix before running Myers to keep the O((N+M)^2) worst
  // case (and memory use) away from long shared prefixes.
  const prefixLen = commonPrefixLength(oldTokens, newTokens, equals);
  const suffixLen = commonSuffixLength(oldTokens, newTokens, equals, prefixLen);

  const oldMid = oldTokens.slice(prefixLen, oldTokens.length - suffixLen);
  const newMid = newTokens.slice(prefixLen, newTokens.length - suffixLen);

  // Guardrail: Myers stores the full trace to reconstruct an edit script.
  // For extremely long formulas, fall back to a simple delete+insert diff for
  // the middle section to keep memory bounded.
  const MAX_MYERS_TOKENS = 2048;
  const midEdits =
    oldMid.length + newMid.length > MAX_MYERS_TOKENS
      ? [
          ...oldMid.map((token) => ({ type: "delete", token })),
          ...newMid.map((token) => ({ type: "insert", token })),
        ]
      : myersDiff(oldMid, newMid, equals);

  /** @type {Array<{ type: DiffOpType, token: Token }>} */
  const edits = [];
  for (const token of oldTokens.slice(0, prefixLen)) edits.push({ type: "equal", token });
  edits.push(...midEdits);
  for (const token of oldTokens.slice(oldTokens.length - suffixLen)) edits.push({ type: "equal", token });

  const ops = coalesceEdits(edits);

  return { equal, ops };
}
