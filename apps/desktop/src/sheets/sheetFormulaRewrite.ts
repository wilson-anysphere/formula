import { rewriteSheetNamesInFormula } from "../workbook/formulaRewrite";

type DocumentControllerLike = {
  setCellInputs?: (
    inputs: ReadonlyArray<{ sheetId: string; row: number; col: number; value: unknown; formula: string | null }>,
    options?: { source?: string; label?: string },
  ) => void;
  // The real DocumentController has a `model` property with a sparse sheet map.
  // This helper intentionally uses an `any` view so it can operate on the runtime
  // JS implementation without importing internal types.
  model?: any;
};

function parseRowColKey(key: string): { row: number; col: number } | null {
  const comma = key.indexOf(",");
  if (comma === -1) return null;
  const row = Number(key.slice(0, comma));
  const col = Number(key.slice(comma + 1));
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

function collectFormulaEdits(params: {
  doc: DocumentControllerLike;
  rewrite: (formula: string) => string;
}): Array<{ sheetId: string; row: number; col: number; value: null; formula: string }> {
  const { doc, rewrite } = params;
  /** @type {Array<{ sheetId: string; row: number; col: number; value: null; formula: string }>} */
  const edits: Array<{ sheetId: string; row: number; col: number; value: null; formula: string }> = [];

  const sheets: Map<string, any> | undefined = doc?.model?.sheets;
  if (!sheets || typeof sheets.entries !== "function") return edits;

  for (const [sheetId, sheet] of sheets.entries()) {
    const cells: Map<string, any> | undefined = sheet?.cells;
    if (!cells || typeof cells.entries !== "function") continue;

    for (const [key, cell] of cells.entries()) {
      const formula = cell?.formula;
      if (formula == null) continue;
      const next = rewrite(String(formula));
      if (next === formula) continue;
      const coord = parseRowColKey(String(key));
      if (!coord) continue;
      edits.push({ sheetId, row: coord.row, col: coord.col, value: null, formula: next });
    }
  }

  return edits;
}

export function rewriteDocumentFormulasForSheetRename(
  doc: DocumentControllerLike,
  oldName: string,
  newName: string,
): void {
  const edits = collectFormulaEdits({
    doc,
    rewrite: (formula) => rewriteSheetNamesInFormula(formula, oldName, newName),
  });
  if (edits.length === 0) return;
  if (typeof doc.setCellInputs !== "function") return;

  doc.setCellInputs(edits, { label: "Rename Sheet", source: "sheetRename" });
}

export function rewriteDocumentFormulasForSheetDelete(doc: DocumentControllerLike, deletedName: string): void {
  const edits = collectFormulaEdits({
    doc,
    rewrite: (formula) => rewriteDeletedSheetReferencesInFormula(formula, deletedName),
  });
  if (edits.length === 0) return;
  if (typeof doc.setCellInputs !== "function") return;

  doc.setCellInputs(edits, { label: "Delete Sheet", source: "sheetDelete" });
}

function splitWorkbookPrefix(sheetSpec: string): { workbookPrefix: string | null; remainder: string } {
  if (!sheetSpec.startsWith("[")) return { workbookPrefix: null, remainder: sheetSpec };
  const closeIdx = sheetSpec.indexOf("]");
  if (closeIdx === -1) return { workbookPrefix: null, remainder: sheetSpec };
  return {
    workbookPrefix: sheetSpec.slice(0, closeIdx + 1),
    remainder: sheetSpec.slice(closeIdx + 1),
  };
}

function split3d(remainder: string): [string, string | null] {
  const idx = remainder.indexOf(":");
  if (idx === -1) return [remainder, null];
  return [remainder.slice(0, idx), remainder.slice(idx + 1)];
}

function sheetSpecMatchesDeletedSheet(sheetSpec: string, deletedName: string): boolean {
  const deletedCi = deletedName.trim().toLowerCase();
  if (!deletedCi) return false;

  const { remainder } = splitWorkbookPrefix(sheetSpec);
  const [start, end] = split3d(remainder);
  const startCi = start.trim().toLowerCase();
  const endCi = end?.trim().toLowerCase() ?? null;
  return startCi === deletedCi || (endCi != null && endCi === deletedCi);
}

function parseQuotedSheetSpec(formula: string, startIndex: number): { nextIndex: number; sheetSpec: string } | null {
  if (formula[startIndex] !== "'") return null;

  let i = startIndex + 1;
  const content: string[] = [];

  while (i < formula.length) {
    const ch = formula[i];
    if (ch === "'") {
      if (formula[i + 1] === "'") {
        content.push("'");
        i += 2;
        continue;
      }
      i += 1;
      break;
    }
    content.push(ch);
    i += 1;
  }

  if (formula[i] !== "!") return null;
  return { nextIndex: i + 1, sheetSpec: content.join("") };
}

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiAlphaNum(ch: string): boolean {
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function parseUnquotedSheetSpec(formula: string, startIndex: number): { nextIndex: number; sheetSpec: string } | null {
  const first = formula[startIndex];
  if (!first || !(isAsciiLetter(first) || first === "_")) return null;

  let i = startIndex;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === "!") {
      return { nextIndex: i + 1, sheetSpec: formula.slice(startIndex, i) };
    }
    if (isAsciiAlphaNum(ch) || ch === "_" || ch === "." || ch === ":") {
      i += 1;
      continue;
    }
    break;
  }

  return null;
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function isIdentifierStart(ch: string): boolean {
  return isAsciiLetter(ch) || ch === "_";
}

function isIdentifierPart(ch: string): boolean {
  return isIdentifierStart(ch) || isAsciiDigit(ch) || ch === ".";
}

function isReferenceDelimiter(ch: string): boolean {
  if (isWhitespace(ch)) return true;
  return ch === "," || ch === ")" || ch === "(" || "+-*/^&=><;".includes(ch);
}

function tryReadCellRef(input: string, start: number): number | null {
  let i = start;
  if (input[i] === "$") i += 1;

  const colStart = i;
  while (i < input.length) {
    const ch = input[i];
    if (!isAsciiLetter(ch)) break;
    i += 1;
  }
  if (i === colStart) return null;
  const colLen = i - colStart;
  if (colLen > 3) return null;

  if (input[i] === "$") i += 1;

  const rowStart = i;
  while (isAsciiDigit(input[i] ?? "")) i += 1;
  if (i === rowStart) return null;

  return i;
}

function tryReadIdentifier(input: string, start: number): number | null {
  if (!isIdentifierStart(input[start] ?? "")) return null;
  let i = start + 1;
  while (i < input.length && isIdentifierPart(input[i]!)) i += 1;
  return i;
}

function tryReadErrorCode(input: string, start: number): number | null {
  if (input[start] !== "#") return null;
  let i = start + 1;
  while (i < input.length) {
    const ch = input[i];
    if (isReferenceDelimiter(ch)) break;
    i += 1;
  }
  return i === start + 1 ? null : i;
}

function skipReferenceTail(formula: string, startIndex: number): number {
  let i = startIndex;
  while (i < formula.length && isWhitespace(formula[i]!)) i += 1;
  if (i >= formula.length) return i;

  const cellEnd = tryReadCellRef(formula, i);
  if (cellEnd != null) {
    i = cellEnd;
    if (formula[i] === "#") i += 1;
    if (formula[i] === ":") {
      const secondEnd = tryReadCellRef(formula, i + 1);
      if (secondEnd != null) {
        i = secondEnd;
        if (formula[i] === "#") i += 1;
      }
    }
    return i;
  }

  const identEnd = tryReadIdentifier(formula, i);
  if (identEnd != null) return identEnd;

  const errEnd = tryReadErrorCode(formula, i);
  if (errEnd != null) return errEnd;

  // Fallback: consume until we hit a delimiter, another sheet separator, or a string.
  while (i < formula.length) {
    const ch = formula[i]!;
    if (isReferenceDelimiter(ch) || ch === "!" || ch === '"') break;
    i += 1;
  }
  return i;
}

export function rewriteDeletedSheetReferencesInFormula(formula: string, deletedSheetName: string): string {
  const deletedCi = deletedSheetName.trim().toLowerCase();
  if (!deletedCi) return formula;

  const out: string[] = [];
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i]!;

    if (inString) {
      out.push(ch);
      if (ch === '"') {
        if (formula[i + 1] === '"') {
          out.push('"');
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }

    if (ch === '"') {
      inString = true;
      out.push(ch);
      i += 1;
      continue;
    }

    if (ch === "'") {
      const parsed = parseQuotedSheetSpec(formula, i);
      if (parsed) {
        const { sheetSpec, nextIndex } = parsed;
        if (sheetSpecMatchesDeletedSheet(sheetSpec, deletedSheetName)) {
          const refEnd = skipReferenceTail(formula, nextIndex);
          out.push("#REF!");
          i = refEnd;
          continue;
        }
        out.push(formula.slice(i, nextIndex));
        i = nextIndex;
        continue;
      }
    }

    const parsedUnquoted = parseUnquotedSheetSpec(formula, i);
    if (parsedUnquoted) {
      const { sheetSpec, nextIndex } = parsedUnquoted;
      if (sheetSpecMatchesDeletedSheet(sheetSpec, deletedSheetName)) {
        const refEnd = skipReferenceTail(formula, nextIndex);
        out.push("#REF!");
        i = refEnd;
        continue;
      }
      out.push(formula.slice(i, nextIndex));
      i = nextIndex;
      continue;
    }

    out.push(ch);
    i += 1;
  }

  return out.join("");
}
