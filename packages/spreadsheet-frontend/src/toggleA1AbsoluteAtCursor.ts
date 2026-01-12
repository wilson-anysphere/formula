import { extractFormulaReferences } from "./formulaReferences.ts";

export type ToggleA1AbsoluteAtCursorResult = {
  text: string;
  cursorStart: number;
  cursorEnd: number;
};

type ParsedCellRef = {
  colAbs: boolean;
  col: string;
  rowAbs: boolean;
  row: string;
};

type ParsedA1ReferenceToken = {
  /** Full token text (e.g. "A1", "Sheet1!A1:B2"). */
  text: string;
  /** Original prefix (including trailing "!") or "" when unqualified. */
  sheetPrefix: string;
  start: ParsedCellRef;
  end: ParsedCellRef | null;
};

export function toggleA1AbsoluteAtCursor(
  formula: string,
  cursorStart: number,
  cursorEnd: number
): ToggleA1AbsoluteAtCursorResult | null {
  const { references, activeIndex } = extractFormulaReferences(formula, cursorStart, cursorEnd);
  if (activeIndex == null) return null;
  const active = references[activeIndex];
  if (!active) return null;

  const oldTokenText = active.text;
  const parsed = parseA1ReferenceToken(oldTokenText);
  if (!parsed) return null;

  const nextTokenText = toggleParsedReference(parsed);
  if (!nextTokenText) return null;

  const prefix = formula.slice(0, active.start);
  const suffix = formula.slice(active.end);
  const nextFormula = prefix + nextTokenText + suffix;

  // Excel-style UX: after toggling, select the entire updated reference token so repeated
  // F4 presses keep cycling the same token (and range-drag insertion can replace it).
  const tokenStart = active.start;
  const tokenEnd = active.start + nextTokenText.length;
  let nextCursorStart = tokenStart;
  let nextCursorEnd = tokenEnd;
  if (cursorStart > cursorEnd) {
    nextCursorStart = tokenEnd;
    nextCursorEnd = tokenStart;
  }

  // Clamp to the output string (defensive).
  nextCursorStart = Math.max(0, Math.min(nextCursorStart, nextFormula.length));
  nextCursorEnd = Math.max(0, Math.min(nextCursorEnd, nextFormula.length));

  return { text: nextFormula, cursorStart: nextCursorStart, cursorEnd: nextCursorEnd };
}

function parseA1ReferenceToken(text: string): ParsedA1ReferenceToken | null {
  const { sheetPrefix, refText } = splitSheetPrefix(text);

  const first = parseCellRefAt(refText, 0);
  if (!first) return null;

  const afterFirst = refText.slice(first.localEnd);
  if (!afterFirst) {
    return { text, sheetPrefix, start: first.cell, end: null };
  }

  if (!afterFirst.startsWith(":")) return null;
  const second = parseCellRefAt(afterFirst, 1);
  if (!second) return null;
  if (second.localEnd !== afterFirst.length) return null;

  return { text, sheetPrefix, start: first.cell, end: second.cell };
}

function splitSheetPrefix(text: string): { sheetPrefix: string; refText: string } {
  if (text.startsWith("'")) {
    // Excel escapes apostrophes inside sheet names using doubled quotes: ''.
    let i = 1;
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          i += 2;
          continue;
        }
        if (text[i + 1] === "!") {
          const end = i + 2;
          return { sheetPrefix: text.slice(0, end), refText: text.slice(end) };
        }
        break;
      }
      i += 1;
    }
    return { sheetPrefix: "", refText: text };
  }

  const bang = text.indexOf("!");
  if (bang === -1) return { sheetPrefix: "", refText: text };

  const end = bang + 1;
  return { sheetPrefix: text.slice(0, end), refText: text.slice(end) };
}

function parseCellRefAt(
  refText: string,
  localStart: number
): { cell: ParsedCellRef; localEnd: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?([0-9]+)/.exec(refText.slice(localStart));
  if (!match) return null;
  const full = match[0] ?? "";
  const colLetters = match[1] ?? "";
  const rowDigits = match[2] ?? "";
  if (!full || !colLetters || !rowDigits) return null;

  const start = localStart;
  const end = localStart + full.length;
  const cellText = full;

  const colAbs = cellText.startsWith("$");
  const rowAbs = (() => {
    const withoutLeading = colAbs ? cellText.slice(1) : cellText;
    return withoutLeading.includes("$");
  })();

  // Preserve the original casing for column letters (as typed).
  const col = colLetters;
  const row = rowDigits;

  const cell: ParsedCellRef = { colAbs, col, rowAbs, row };
  return { cell, localEnd: end };
}

function toggleParsedReference(parsed: ParsedA1ReferenceToken): string | null {
  const start = toggleCellRef(parsed.start);
  if (!start) return null;

  if (!parsed.end) return parsed.sheetPrefix + start;

  const end = toggleCellRef(parsed.end);
  if (!end) return null;

  return parsed.sheetPrefix + start + ":" + end;
}

function toggleCellRef(cell: ParsedCellRef): string | null {
  const next = nextAbsoluteMode(cell.colAbs, cell.rowAbs);
  const colAbs = next.colAbs;
  const rowAbs = next.rowAbs;
  return `${colAbs ? "$" : ""}${cell.col}${rowAbs ? "$" : ""}${cell.row}`;
}

function nextAbsoluteMode(colAbs: boolean, rowAbs: boolean): { colAbs: boolean; rowAbs: boolean } {
  // Excel cycle:
  //   A1 → $A$1 → A$1 → $A1 → A1
  if (!colAbs && !rowAbs) return { colAbs: true, rowAbs: true };
  if (colAbs && rowAbs) return { colAbs: false, rowAbs: true };
  if (!colAbs && rowAbs) return { colAbs: true, rowAbs: false };
  // colAbs && !rowAbs
  return { colAbs: false, rowAbs: false };
}
