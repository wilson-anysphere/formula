function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Convert an Excel-style wildcard pattern to a RegExp.
 *
 * Excel wildcards:
 * - `*` matches any sequence of characters (including empty)
 * - `?` matches any single character
 * - `~` escapes the following character (e.g. `~*` matches literal `*`)
 */
export function excelWildcardToRegExp(
  pattern,
  {
    matchCase = false,
    matchEntireCell = false,
    useWildcards = true,
    global = false,
  } = {},
) {
  if (pattern == null) {
    throw new TypeError("excelWildcardToRegExp: pattern is required");
  }

  const input = String(pattern);
  let source = "";

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];

    if (useWildcards && ch === "~") {
      const next = input[i + 1];
      if (next == null) {
        source += escapeRegExp("~");
      } else {
        source += escapeRegExp(next);
        i++;
      }
      continue;
    }

    if (useWildcards && ch === "*") {
      source += "[\\s\\S]*";
      continue;
    }

    if (useWildcards && ch === "?") {
      source += "[\\s\\S]";
      continue;
    }

    source += escapeRegExp(ch);
  }

  if (matchEntireCell) {
    source = `^${source}$`;
  }

  let flags = "";
  if (global) flags += "g";
  if (!matchCase) flags += "i";

  return new RegExp(source, flags);
}

export function excelWildcardTest(text, pattern, options) {
  const re = excelWildcardToRegExp(pattern, options);
  return re.test(String(text));
}
