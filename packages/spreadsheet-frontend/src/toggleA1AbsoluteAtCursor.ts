import { extractFormulaReferences } from "./formulaReferences";

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
  /** 0-based start offset within the full reference token text. */
  start: number;
};

type ParsedA1ReferenceToken = {
  /** Full token text (e.g. "A1", "Sheet1!A1:B2"). */
  text: string;
  /** Original prefix (including trailing "!") or "" when unqualified. */
  sheetPrefix: string;
  start: ParsedCellRef;
  end: ParsedCellRef | null;
};

type OffsetOp = { pos: number; kind: "insert" | "remove" };

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

  const toggled = toggleParsedReference(parsed);
  if (!toggled) return null;

  const prefix = formula.slice(0, active.start);
  const suffix = formula.slice(active.end);
  const nextFormula = prefix + toggled.text + suffix;

  // Preserve token-wide selection when the full reference token is selected.
  const start = cursorStart;
  const end = cursorEnd;
  const selMin = Math.min(start, end);
  const selMax = Math.max(start, end);
  const selectionCoversToken = selMin !== selMax && selMin === active.start && selMax === active.end;

  const delta = toggled.text.length - oldTokenText.length;

  const mapCursor = (pos: number): number => {
    // Before/after token: shift by delta.
    if (pos < active.start) return pos;
    if (pos > active.end) return pos + delta;

    const oldOffset = pos - active.start;
    const newOffset = mapOffsetWithinToken(oldOffset, toggled.ops, toggled.text.length);
    return active.start + newOffset;
  };

  let nextCursorStart: number;
  let nextCursorEnd: number;
  if (selectionCoversToken) {
    const tokenStart = active.start;
    const tokenEnd = active.start + toggled.text.length;
    if (start <= end) {
      nextCursorStart = tokenStart;
      nextCursorEnd = tokenEnd;
    } else {
      nextCursorStart = tokenEnd;
      nextCursorEnd = tokenStart;
    }
  } else {
    nextCursorStart = mapCursor(start);
    nextCursorEnd = mapCursor(end);
  }

  // Clamp to the output string (defensive).
  nextCursorStart = Math.max(0, Math.min(nextCursorStart, nextFormula.length));
  nextCursorEnd = Math.max(0, Math.min(nextCursorEnd, nextFormula.length));

  return { text: nextFormula, cursorStart: nextCursorStart, cursorEnd: nextCursorEnd };
}

function mapOffsetWithinToken(oldOffset: number, ops: readonly OffsetOp[], newTokenLen: number): number {
  let next = oldOffset;
  for (const op of ops) {
    if (op.kind === "insert") {
      if (oldOffset >= op.pos) next += 1;
      continue;
    }
    // remove
    if (oldOffset > op.pos) next -= 1;
  }
  return Math.max(0, Math.min(next, newTokenLen));
}

function parseA1ReferenceToken(text: string): ParsedA1ReferenceToken | null {
  const { sheetPrefix, refText } = splitSheetPrefix(text);
  const prefixLen = sheetPrefix.length;

  const first = parseCellRefAt(refText, 0, prefixLen);
  if (!first) return null;

  const afterFirst = refText.slice(first.localEnd);
  if (!afterFirst) {
    return { text, sheetPrefix, start: first.cell, end: null };
  }

  if (!afterFirst.startsWith(":")) return null;
  // `afterFirst` starts at `refText[first.localEnd]`, which corresponds to
  // `prefixLen + first.localEnd` within the original token.
  const second = parseCellRefAt(afterFirst, 1, prefixLen + first.localEnd);
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

  const candidate = text.slice(0, bang);
  // Match the tokenizer: unquoted sheet names must be identifier-ish.
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(candidate)) return { sheetPrefix: "", refText: text };

  const end = bang + 1;
  return { sheetPrefix: text.slice(0, end), refText: text.slice(end) };
}

function parseCellRefAt(
  refText: string,
  localStart: number,
  absoluteStartWithinToken: number
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

  const absoluteStart = absoluteStartWithinToken + start;

  const cell: ParsedCellRef = { colAbs, col, rowAbs, row, start: absoluteStart };
  return { cell, localEnd: end };
}

function toggleParsedReference(parsed: ParsedA1ReferenceToken): { text: string; ops: OffsetOp[] } | null {
  const ops: OffsetOp[] = [];

  const start = toggleCellRef(parsed.start, ops);
  if (!start) return null;

  if (!parsed.end) {
    return { text: parsed.sheetPrefix + start.text, ops };
  }

  const end = toggleCellRef(parsed.end, ops);
  if (!end) return null;

  return { text: parsed.sheetPrefix + start.text + ":" + end.text, ops };
}

function toggleCellRef(cell: ParsedCellRef, ops: OffsetOp[]): { text: string } | null {
  const next = nextAbsoluteMode(cell.colAbs, cell.rowAbs);
  const colAbs = next.colAbs;
  const rowAbs = next.rowAbs;

  // Compute mapping ops (relative to the original token text).
  // Leading `$` toggles at the cell start.
  if (cell.colAbs !== colAbs) {
    ops.push({ pos: cell.start, kind: cell.colAbs ? "remove" : "insert" });
  }

  // Row `$` toggles after the column letters (and any leading `$`).
  const oldLeadingLen = cell.colAbs ? 1 : 0;
  const rowDollarPos = cell.start + oldLeadingLen + cell.col.length;
  const rowDigitsStart = rowDollarPos + (cell.rowAbs ? 1 : 0);

  if (cell.rowAbs !== rowAbs) {
    ops.push({ pos: cell.rowAbs ? rowDollarPos : rowDigitsStart, kind: cell.rowAbs ? "remove" : "insert" });
  }

  const text = `${colAbs ? "$" : ""}${cell.col}${rowAbs ? "$" : ""}${cell.row}`;
  return { text };
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
