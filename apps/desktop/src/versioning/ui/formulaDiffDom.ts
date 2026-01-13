import { diffFormulaToRenderOps, isEffectivelyEmptyFormula } from "./formulaDiffRender.ts";

export type RenderFormulaDiffDomOptions = {
  /**
   * Optional `data-testid` for callers/tests.
   */
  testid?: string;
  /**
   * Optional label rendered above the diff (defaults to "Diff").
   */
  label?: string;
};

function joinClassName(...classes: Array<string | undefined | null | false>): string {
  return classes.filter(Boolean).join(" ");
}

/**
 * Create a small DOM subtree that renders a token diff between two formulas.
 *
 * This is used by non-React UIs (e.g. the minimal conflict controllers) while
 * sharing the same token formatting + spacing heuristics as {@link FormulaDiffView}.
 */
export function renderFormulaDiffDom(
  before: string | null,
  after: string | null,
  opts: RenderFormulaDiffDomOptions = {},
): HTMLElement {
  const root = document.createElement("div");
  root.className = "conflict-dialog__formula-diff";
  if (opts.testid) root.dataset.testid = opts.testid;

  const label = document.createElement("div");
  label.className = "conflict-dialog__formula-diff-label";
  label.textContent = opts.label ?? "Diff";
  root.appendChild(label);

  const code = document.createElement("code");
  code.className = joinClassName("formula-diff-view");

  const emptyBefore = isEffectivelyEmptyFormula(before);
  const emptyAfter = isEffectivelyEmptyFormula(after);
  if (emptyBefore && emptyAfter) {
    code.classList.add("formula-diff-view--empty");
    const marker = document.createElement("span");
    marker.className = "formula-diff-empty-marker";
    marker.textContent = "∅";
    code.appendChild(marker);
    root.appendChild(code);
    return root;
  }

  const ops = diffFormulaToRenderOps(before, after);
  for (const op of ops) {
    const span = document.createElement("span");
    span.className =
      op.type === "insert"
        ? "formula-diff-op formula-diff-op--insert"
        : op.type === "delete"
          ? "formula-diff-op formula-diff-op--delete"
          : "formula-diff-op formula-diff-op--equal";
    span.textContent = op.text;
    code.appendChild(span);
  }

  root.appendChild(code);
  return root;
}
