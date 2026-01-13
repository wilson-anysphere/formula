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

function tokenToText(token: Token): string {
  if (token.type === "string") return `"${escapeExcelString(token.value)}"`;
  return token.value;
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

  return (
    <code className={joinClassName("formula-diff-view", className)}>
      {ops.map((op, opIndex) => {
        const opClassName =
          op.type === "insert"
            ? "formula-diff-op formula-diff-op--insert"
            : op.type === "delete"
              ? "formula-diff-op formula-diff-op--delete"
              : "formula-diff-op formula-diff-op--equal";

        return (
          <span key={opIndex} className={opClassName}>
            {op.tokens.map((token, tokenIndex) => (
              <React.Fragment key={tokenIndex}>{tokenToText(token)}</React.Fragment>
            ))}
          </span>
        );
      })}
    </code>
  );
}

