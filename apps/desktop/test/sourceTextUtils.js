export function skipStringLiteral(source, start) {
  const quote = source[start];
  if (quote !== "'" && quote !== '"' && quote !== "`") return start + 1;
  if (quote === "`") return skipTemplateLiteral(source, start);

  let i = start + 1;
  while (i < source.length) {
    const ch = source[i];
    if (ch === "\\") {
      i += 2;
      continue;
    }
    if (ch === quote) return i + 1;
    i += 1;
  }
  return source.length;
}

function skipTemplateLiteral(source, start) {
  // Template literals can contain nested `${ ... }` expressions (including other template literals).
  // This is a lightweight best-effort skipper so source-scanning guardrails don't get confused
  // by backticks appearing inside `${}`.
  let i = start + 1;
  while (i < source.length) {
    const ch = source[i];
    if (ch === "\\") {
      // Escape sequence inside the template quasi.
      i += 2;
      continue;
    }

    if (ch === "`") return i + 1;

    if (ch === "$" && source[i + 1] === "{") {
      i += 2; // consume `${`
      let depth = 1;
      for (; i < source.length; i += 1) {
        const c = source[i];

        if (c === "'" || c === '"' || c === "`") {
          i = skipStringLiteral(source, i) - 1;
          continue;
        }

        // Skip comments inside the expression so braces inside comments don't affect depth.
        if (c === "/" && source[i + 1] === "/") {
          i += 2;
          while (i < source.length && source[i] !== "\n") i += 1;
          i -= 1;
          continue;
        }

        if (c === "/" && source[i + 1] === "*") {
          i += 2;
          while (i < source.length) {
            if (source[i] === "*" && source[i + 1] === "/") {
              i += 1;
              break;
            }
            i += 1;
          }
          continue;
        }

        if (c === "/" && isRegexLiteralStart(source, i)) {
          i = skipRegexLiteral(source, i) - 1;
          continue;
        }

        if (c === "{") depth += 1;
        else if (c === "}") {
          depth -= 1;
          if (depth === 0) break;
        }
      }

      if (i >= source.length) return source.length;
      // `i` is at the closing `}`; continue scanning the template quasi.
    }

    i += 1;
  }
  return source.length;
}

function isRegexLiteralStart(source, start) {
  // Best-effort detection of regex literals vs division.
  // This is intentionally heuristic (not a full JS lexer), but handles the patterns used in
  // our guardrail tests (e.g. `.split(/[/\\\\]/)` or `a / /re/.test(x)`).
  if (source[start] !== "/") return false;
  const next = source[start + 1];
  if (next == null) return false;
  // `//` and `/*` are always comments (regex bodies cannot start with `/` or `*`).
  if (next === "/" || next === "*") return false;

  let i = start - 1;
  while (i >= 0 && /\s/.test(source[i])) i -= 1;
  if (i < 0) return true;
  const prev = source[i];

  // Characters that can precede an expression, where a regex literal is valid.
  if ("([{:;,=!?&|+-*%^~<>/".includes(prev)) return true;

  // Keywords like `return /.../` are also valid.
  if (/[A-Za-z_$]/.test(prev)) {
    let j = i;
    while (j >= 0 && /[A-Za-z0-9_$]/.test(source[j])) j -= 1;
    const word = source.slice(j + 1, i + 1);
    if (
      word === "return" ||
      word === "throw" ||
      word === "case" ||
      word === "delete" ||
      word === "typeof" ||
      word === "void" ||
      word === "in" ||
      word === "of" ||
      word === "instanceof"
    ) {
      return true;
    }
  }

  return false;
}

function skipRegexLiteral(source, start) {
  // Assumes `source[start] === '/'` and `isRegexLiteralStart(source, start) === true`.
  let inCharClass = false;
  let escaped = false;

  for (let i = start + 1; i < source.length; i += 1) {
    const ch = source[i];
    if (escaped) {
      escaped = false;
      continue;
    }

    if (ch === "\\") {
      escaped = true;
      continue;
    }

    if (ch === "[") {
      inCharClass = true;
      continue;
    }

    if (ch === "]" && inCharClass) {
      inCharClass = false;
      continue;
    }

    if (ch === "/" && !inCharClass) {
      // Consume regex flags.
      let j = i + 1;
      while (j < source.length && /[A-Za-z]/.test(source[j])) j += 1;
      return j;
    }
  }

  return source.length;
}

export function stripComments(source) {
  // Remove JS comments without accidentally stripping `https://...` inside string literals.
  // This is intentionally lightweight: it's not a full parser, but is sufficient for guardrail
  // matching in `main.ts` and avoids treating commented-out wiring as valid.
  let out = "";
  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];
    const prev = i > 0 ? source[i - 1] : "";

    if (ch === "'" || ch === '"' || ch === "`") {
      const end = skipStringLiteral(source, i);
      out += source.slice(i, end);
      i = end - 1;
      continue;
    }

    // Treat `//` and `/*` as comments unless they are preceded by a backslash. This avoids
    // accidentally stripping the `\/` at the end of a regex literal like `/foo\//`, which would
    // otherwise look like the start of a line comment to this lightweight scanner.
    if (ch === "/" && source[i + 1] === "/" && prev !== "\\") {
      // Line comment.
      i += 2;
      while (i < source.length && source[i] !== "\n") i += 1;
      if (i < source.length) out += "\n";
      continue;
    }

    if (ch === "/" && source[i + 1] === "*" && prev !== "\\") {
      // Block comment (preserve newlines so we don't accidentally join tokens across lines).
      i += 2;
      while (i < source.length) {
        const next = source[i];
        if (next === "\n") out += "\n";
        if (next === "*" && source[i + 1] === "/") {
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }

    if (ch === "/" && isRegexLiteralStart(source, i)) {
      const end = skipRegexLiteral(source, i);
      out += source.slice(i, end);
      i = end - 1;
      continue;
    }

    out += ch;
  }
  return out;
}

