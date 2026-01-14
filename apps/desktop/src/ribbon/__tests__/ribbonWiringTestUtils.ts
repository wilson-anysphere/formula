export type OverrideKey = "commandOverrides" | "toggleOverrides";

export function extractObjectLiteral(source: string, key: OverrideKey): string | null {
  const idx = source.indexOf(`${key}:`);
  if (idx === -1) return null;
  const braceStart = source.indexOf("{", idx);
  if (braceStart === -1) return null;

  let depth = 0;
  let inString: '"' | "'" | "`" | null = null;
  let inLineComment = false;
  let inBlockComment = false;

  for (let i = braceStart; i < source.length; i += 1) {
    const ch = source[i];
    const next = source[i + 1];

    if (inLineComment) {
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (ch === "*" && next === "/") {
        inBlockComment = false;
        i += 1;
      }
      continue;
    }

    if (inString) {
      if (ch === "\\") {
        i += 1;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }

    if (ch === "/" && next === "/") {
      inLineComment = true;
      i += 1;
      continue;
    }

    if (ch === "/" && next === "*") {
      inBlockComment = true;
      i += 1;
      continue;
    }

    if (ch === '"' || ch === "'" || ch === "`") {
      inString = ch;
      continue;
    }

    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(braceStart, i + 1);
    }
  }

  return null;
}

export function extractTopLevelStringKeys(objectText: string): string[] {
  const keys: string[] = [];
  let depth = 0;
  let inLineComment = false;
  let inBlockComment = false;

  const skipWhitespace = (idx: number): number => {
    while (idx < objectText.length && /\s/.test(objectText[idx])) idx += 1;
    return idx;
  };

  for (let i = 0; i < objectText.length; i += 1) {
    const ch = objectText[i];
    const next = objectText[i + 1];

    if (inLineComment) {
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (ch === "*" && next === "/") {
        inBlockComment = false;
        i += 1;
      }
      continue;
    }

    if (ch === "/" && next === "/") {
      inLineComment = true;
      i += 1;
      continue;
    }

    if (ch === "/" && next === "*") {
      inBlockComment = true;
      i += 1;
      continue;
    }

    if (ch === "{") {
      depth += 1;
      continue;
    }

    if (ch === "}") {
      depth -= 1;
      continue;
    }

    if (depth !== 1) continue;
    if (ch !== '"' && ch !== "'") continue;

    const quote = ch;
    let j = i + 1;
    let value = "";

    for (; j < objectText.length; j += 1) {
      const c = objectText[j];
      if (c === "\\") {
        value += objectText[j + 1] ?? "";
        j += 1;
        continue;
      }
      if (c === quote) break;
      value += c;
    }

    if (j >= objectText.length) break;

    const k = skipWhitespace(j + 1);
    if (objectText[k] === ":") {
      keys.push(value);
    }

    i = j;
  }

  return keys;
}

