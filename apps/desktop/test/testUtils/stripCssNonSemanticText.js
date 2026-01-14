/**
 * Remove CSS comments, string literals, and url(...) bodies so we don't flag
 * `border-radius:` text inside non-semantic content (e.g. `content: "border-radius: 4px"` or
 * data URIs).
 *
 * This is intentionally lightweight (not a full CSS parser), but it preserves newline
 * characters so line numbers remain accurate.
 *
 * @param {string} css
 */
export function stripCssNonSemanticText(css) {
  let out = String(css);

  // Block comments (preserve newlines).
  out = out.replace(/\/\*[\s\S]*?\*\//g, (comment) => comment.replace(/[^\n]/g, " "));

  // Quoted strings (handles escapes; preserves newlines in escaped multi-line strings).
  out = out.replace(/"(?:\\.|[^"\\])*"/g, (str) => str.replace(/[^\n]/g, " "));
  out = out.replace(/'(?:\\.|[^'\\])*'/g, (str) => str.replace(/[^\n]/g, " "));

  // Strip url(...) bodies while preserving newlines so line numbers stay stable.
  let idx = 0;
  let result = "";
  while (idx < out.length) {
    const m = /\burl\s*\(/gi.exec(out.slice(idx));
    if (!m) {
      result += out.slice(idx);
      break;
    }

    const start = idx + (m.index ?? 0);
    result += out.slice(idx, start);

    const openParen = out.indexOf("(", start);
    if (openParen === -1) {
      result += out.slice(start);
      break;
    }

    // Copy the `url(` prefix.
    result += out.slice(start, openParen + 1);

    let i = openParen + 1;
    let depth = 1;
    /** @type {string | null} */
    let quote = null;
    while (i < out.length && depth > 0) {
      const ch = out[i];
      const next = i + 1 < out.length ? out[i + 1] : "";

      if (quote) {
        if (ch === "\\") {
          result += ch === "\n" ? "\n" : " ";
          if (next) {
            result += next === "\n" ? "\n" : " ";
            i += 2;
            continue;
          }
          i += 1;
          continue;
        }

        if (ch === quote) quote = null;
        result += ch === "\n" ? "\n" : " ";
        i += 1;
        continue;
      }

      if (ch === '"' || ch === "'") {
        quote = ch;
        result += " ";
        i += 1;
        continue;
      }

      if (ch === "(") {
        depth += 1;
        result += " ";
        i += 1;
        continue;
      }

      if (ch === ")") {
        depth -= 1;
        if (depth === 0) {
          result += ")";
          i += 1;
          break;
        }
        result += " ";
        i += 1;
        continue;
      }

      result += ch === "\n" ? "\n" : " ";
      i += 1;
    }

    idx = i;
  }

  return result;
}

