import { tokenizeFormula } from "./tokenizeFormula.js";
import type { HighlightSpan } from "./types.js";

export type { HighlightSpan } from "./types.js";

export function highlightFormula(input: string): HighlightSpan[] {
  return tokenizeFormula(input).map((token) => ({
    kind: token.type,
    text: token.text,
    start: token.start,
    end: token.end,
  }));
}

function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export function highlightFormulaToHtml(input: string): string {
  return highlightFormula(input)
    .map((span) => `<span data-kind="${span.kind}">${escapeHtml(span.text)}</span>`)
    .join("");
}
