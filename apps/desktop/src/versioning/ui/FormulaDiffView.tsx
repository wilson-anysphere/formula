import React, { useMemo } from "react";

import { diffFormulaToRenderOps, isEffectivelyEmptyFormula } from "./formulaDiffRender.js";
import { t } from "../../i18n/index.js";

export type FormulaDiffViewProps = {
  before: string | null;
  after: string | null;
  /**
   * Optional extra class name for the root element.
   */
  className?: string;
};

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
        <span className="formula-diff-empty-marker" aria-label={t("formulaDiff.aria.emptyFormula")}>
          ∅
        </span>
      </code>
    );
  }

  // `diffFormula` is authored in JS (workspace package) so keep the UI boundary
  // typed locally.
  const ops = useMemo(() => diffFormulaToRenderOps(before, after), [before, after]);

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
            {op.text}
          </span>
        );
      })}
    </code>
  );
}
