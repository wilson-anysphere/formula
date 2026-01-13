export type FormulaTokenType =
  | "whitespace"
  | "operator"
  | "punctuation"
  | "number"
  | "string"
  | "function"
  | "identifier"
  | "reference"
  | "error"
  | "unknown";

export type HighlightKind =
  | "whitespace"
  | "operator"
  | "punctuation"
  | "number"
  | "string"
  | "function"
  | "identifier"
  | "reference"
  | "error"
  | "unknown";

export type FormulaToken = {
  type: FormulaTokenType;
  text: string;
  start: number;
  end: number;
};

export type HighlightSpan = {
  kind: HighlightKind;
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
