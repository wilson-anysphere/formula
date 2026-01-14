import { tokenizeFormula } from "./tokenizeFormula.js";
import type { HighlightSpan } from "./types.js";

export function highlightFormula(input: string): HighlightSpan[] {
  return tokenizeFormula(input).map((token) => ({
    kind: token.type,
    text: token.text,
    start: token.start,
    end: token.end,
  }));
}
