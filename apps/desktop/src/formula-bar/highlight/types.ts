import type { FormulaTokenType } from "@formula/spreadsheet-frontend/formula/tokenizeFormula";

export type HighlightSpan = {
  kind: FormulaTokenType;
  text: string;
  start: number;
  end: number;
  /**
   * Optional CSS class applied to the rendered <span>.
   *
   * Used by the WASM-backed editor tooling integration to surface parse errors
   * with an exact span highlight.
   */
  className?: string;
};
