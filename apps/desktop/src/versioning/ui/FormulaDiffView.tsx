import React from "react";

import { diffFormula } from "../index.js";

type Token = { type: string; value: string };
type DiffOpType = "equal" | "insert" | "delete";
type DiffOp = { type: DiffOpType; tokens: Token[] };

export type FormulaDiffViewProps = {
  before: string | null;
  after: string | null;
  /**
   * Optional extra class name for the root element.
   */
  className?: string;
};

function isEffectivelyEmptyFormula(formula: string | null): boolean {
  if (formula == null) return true;
  const trimmed = String(formula).trim();
  if (!trimmed) return true;
  const withoutEquals = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  return !withoutEquals.trim();
}

function escapeExcelString(value: string): string {
  // Excel escapes quotes inside string literals as doubled quotes: "".
  return value.replaceAll(`"`, `""`);
}

function escapeExcelSheetName(value: string): string {
  // Excel escapes apostrophes inside quoted sheet names by doubling: ''.
  return value.replaceAll(`'`, `''`);
}

function isDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isReservedUnquotedSheetName(name: string): boolean {
  const lower = String(name ?? "").toLowerCase();
  return lower === "true" || lower === "false";
}

function looksLikeA1CellReference(name: string): boolean {
  // Mirrors the FormulaBar tokenizer rules: treat names like "A1" as requiring sheet quotes.
  let i = 0;
  let letters = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiLetter(ch)) break;
    if (letters.length >= 3) return false;
    letters += ch;
    i += 1;
  }

  if (letters.length === 0) return false;

  let digits = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isDigit(ch)) break;
    digits += ch;
    i += 1;
  }

  if (digits.length === 0) return false;
  if (i !== name.length) return false;

  const col = letters
    .split("")
    .reduce((acc, c) => acc * 26 + (c.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1), 0);
  return col <= 16_384;
}

function looksLikeR1C1CellReference(name: string): boolean {
  const upper = String(name ?? "").toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;

  let i = 1;
  while (i < upper.length && isDigit(upper[i] ?? "")) i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;

  i += 1;
  while (i < upper.length && isDigit(upper[i] ?? "")) i += 1;
  return i === upper.length;
}

function shouldQuoteSheetName(name: string): boolean {
  const text = String(name ?? "");
  if (!text) return true;
  if (isReservedUnquotedSheetName(text)) return true;
  if (looksLikeA1CellReference(text)) return true;
  if (looksLikeR1C1CellReference(text)) return true;

  // Unquoted sheet names follow identifier-like rules. Be conservative: if we see
  // anything outside a small ASCII set, re-add quotes.
  const first = text[0] ?? "";
  if (!(isAsciiLetter(first) || first === "_")) return true;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? "";
    if (isAsciiLetter(ch) || isDigit(ch) || ch === "_" || ch === "." || ch === "$") continue;
    return true;
  }

  return false;
}

function tokenToText(token: Token, nextToken: Token | null): string {
  if (token.type === "string") return `"${escapeExcelString(token.value)}"`;
  if (token.type === "ident" && nextToken?.type === "op" && nextToken.value === "!" && shouldQuoteSheetName(token.value)) {
    return `'${escapeExcelSheetName(token.value)}'`;
  }
  return token.value;
}

const BINARY_OPS = new Set(["+", "-", "*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>"]);
const NO_SPACE_OPS = new Set([":", "!", "%"]);
const OPENING_PUNCT = new Set(["(", "{", "["]);
const SEPARATOR_PUNCT = new Set([",", ";"]);

function isSeparatorPunct(token: Token | null): boolean {
  return token?.type === "punct" && SEPARATOR_PUNCT.has(token.value);
}

function isOpeningPunct(token: Token | null): boolean {
  return token?.type === "punct" && OPENING_PUNCT.has(token.value);
}

function isBinaryOperator(token: Token | null, prevToken: Token | null): boolean {
  if (!token || token.type !== "op") return false;
  if (!BINARY_OPS.has(token.value)) return false;
  if (NO_SPACE_OPS.has(token.value)) return false;

  // Leading `=` is the formula prefix, not a comparison.
  if (token.value === "=" && !prevToken) return false;

  // Detect unary +/- based on context.
  if (token.value === "+" || token.value === "-") {
    if (!prevToken) return false;
    if (prevToken.type === "op") return false;
    if (isOpeningPunct(prevToken)) return false;
    if (isSeparatorPunct(prevToken)) return false;
  }

  return true;
}

function shouldInsertLeadingSpace(prevPrevToken: Token | null, prevToken: Token | null, currToken: Token): boolean {
  if (!prevToken) return false;

  // Space after argument separators (commas/semicolons).
  if (isSeparatorPunct(prevToken)) return true;

  // Space around binary operators: insert a space both before and after.
  if (isBinaryOperator(currToken, prevToken)) return true;
  if (isBinaryOperator(prevToken, prevPrevToken)) return true;

  return false;
}

function joinClassName(...classes: Array<string | undefined | null | false>): string {
  return classes.filter(Boolean).join(" ");
}

/**
 * Render a token-level diff of two Excel formulas.
 *
 * Note: the underlying diff algorithm is intentionally whitespace-insensitive for
 * rendering/UX purposes (see `diffFormula` docs), so we render the canonical token
 * stream rather than trying to preserve original spacing.
 */
export function FormulaDiffView({ before, after, className }: FormulaDiffViewProps): React.ReactElement {
  const isEmptyBefore = isEffectivelyEmptyFormula(before);
  const isEmptyAfter = isEffectivelyEmptyFormula(after);

  if (isEmptyBefore && isEmptyAfter) {
    return (
      <code className={joinClassName("formula-diff-view", "formula-diff-view--empty", className)}>
        <span className="formula-diff-empty-marker" aria-label="Empty formula">
          ∅
        </span>
      </code>
    );
  }

  // `diffFormula` is authored in JS (workspace package) so keep the UI boundary
  // typed locally.
  const { ops } = diffFormula(before, after) as { equal: boolean; ops: DiffOp[] };

  /** @type {React.ReactNode[]} */
  const renderedOps: React.ReactNode[] = [];
  let prevToken: Token | null = null;
  let prevPrevToken: Token | null = null;

  for (let opIndex = 0; opIndex < ops.length; opIndex += 1) {
    const op = ops[opIndex]!;

    const opClassName =
      op.type === "insert"
        ? "formula-diff-op formula-diff-op--insert"
        : op.type === "delete"
          ? "formula-diff-op formula-diff-op--delete"
          : "formula-diff-op formula-diff-op--equal";

    /** @type {React.ReactNode[]} */
    const children: React.ReactNode[] = [];

    for (let tokenIndex = 0; tokenIndex < op.tokens.length; tokenIndex += 1) {
      const token = op.tokens[tokenIndex]!;
      let nextToken: Token | null = op.tokens[tokenIndex + 1] ?? null;
      if (!nextToken) {
        for (let nextOpIdx = opIndex + 1; nextOpIdx < ops.length; nextOpIdx += 1) {
          const candidate = ops[nextOpIdx]?.tokens?.[0] ?? null;
          if (candidate) {
            nextToken = candidate;
            break;
          }
        }
      }
      if (shouldInsertLeadingSpace(prevPrevToken, prevToken, token)) {
        children.push(" ");
      }
      children.push(tokenToText(token, nextToken));
      prevPrevToken = prevToken;
      prevToken = token;
    }

    renderedOps.push(
      <span key={opIndex} className={opClassName}>
        {children}
      </span>
    );
  }

  return (
    <code className={joinClassName("formula-diff-view", className)}>
      {renderedOps}
    </code>
  );
}
