export function rewriteSheetNamesInFormula(
  formula: string,
  oldName: string,
  newName: string,
): string {
  const out: string[] = [];
  let i = 0;
  let inString = false;

  while (i < formula.length) {
    const ch = formula[i];

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
      out.push('"');
      i += 1;
      continue;
    }

    if (ch === "'") {
      const parsed = parseQuotedSheetSpec(formula, i);
      if (parsed) {
        const { nextIndex, sheetSpec } = parsed;
        const rewritten =
          rewriteSheetSpec(sheetSpec, oldName, newName) ?? quoteSheetSpec(sheetSpec);
        out.push(rewritten, "!");
        i = nextIndex;
        continue;
      }
    }

    const parsedUnquoted = parseUnquotedSheetSpec(formula, i);
    if (parsedUnquoted) {
      const { nextIndex, sheetSpec } = parsedUnquoted;
      out.push(rewriteSheetSpec(sheetSpec, oldName, newName) ?? sheetSpec, "!");
      i = nextIndex;
      continue;
    }

    out.push(ch);
    i += 1;
  }

  return out.join("");
}

function parseQuotedSheetSpec(
  formula: string,
  startIndex: number,
): { nextIndex: number; sheetSpec: string } | null {
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

function parseUnquotedSheetSpec(
  formula: string,
  startIndex: number,
): { nextIndex: number; sheetSpec: string } | null {
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

function rewriteSheetSpec(sheetSpec: string, oldName: string, newName: string): string | null {
  const { workbookPrefix, remainder } = splitWorkbookPrefix(sheetSpec);
  const [start, end] = split3d(remainder);

  const renamedStart = startEquals(start, oldName) ? newName : start;
  const renamedEnd = end && startEquals(end, oldName) ? newName : end;

  if (renamedStart === start && renamedEnd === end) return null;

  return formatSheetReference(workbookPrefix, renamedStart, renamedEnd);
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

function startEquals(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

function quoteSheetSpec(sheetSpec: string): string {
  return `'${sheetSpec.replace(/'/g, "''")}'`;
}

function isValidUnquotedSheetName(name: string): boolean {
  if (!name) return false;
  const first = name[0];
  if (!first || isAsciiDigit(first)) return false;
  if (!(isAsciiLetter(first) || first === "_")) return false;
  for (let i = 1; i < name.length; i += 1) {
    const ch = name[i];
    if (!(isAsciiAlphaNum(ch) || ch === "_")) return false;
  }
  return true;
}

function needsQuotingForSheetReference(name: string): boolean {
  const [start, end] = split3d(name);
  if (end !== null) {
    return !(isValidUnquotedSheetName(start) && isValidUnquotedSheetName(end));
  }
  return !isValidUnquotedSheetName(name);
}

function formatSheetReference(
  workbookPrefix: string | null,
  start: string,
  end: string | null,
): string {
  const content = `${workbookPrefix ?? ""}${start}${end ? `:${end}` : ""}`;
  return needsQuotingForSheetReference(content) ? quoteSheetSpec(content) : content;
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

