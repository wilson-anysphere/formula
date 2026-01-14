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

export function stripCssComments(css) {
  // Strip CSS block comments while preserving string literals and newlines.
  //
  // This keeps "source scanning" guardrail tests high-signal:
  // - commented-out selectors/declarations should not satisfy or fail assertions
  // - we avoid joining tokens across comments by preserving whitespace/newlines
  const text = String(css);
  let out = "";
  /** @type {"code" | "single" | "double" | "comment"} */
  let state = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : "";

    if (state === "comment") {
      if (ch === "*" && next === "/") {
        out += "  ";
        i += 1;
        state = "code";
        continue;
      }
      out += ch === "\n" ? "\n" : " ";
      continue;
    }

    if (state === "code") {
      if (ch === "/" && next === "*") {
        out += "  ";
        i += 1;
        state = "comment";
        continue;
      }

      if (ch === "'") {
        state = "single";
        out += ch;
        continue;
      }

      if (ch === '"') {
        state = "double";
        out += ch;
        continue;
      }

      out += ch;
      continue;
    }

    // String literal (single/double quote): preserve content and escapes.
    out += ch;
    if (ch === "\\") {
      if (next) {
        out += next;
        i += 1;
      }
      continue;
    }

    if (state === "single" && ch === "'") {
      state = "code";
    } else if (state === "double" && ch === '"') {
      state = "code";
    }
  }

  return out;
}

export function stripHtmlComments(html) {
  // Strip HTML comments (`<!-- ... -->`) while preserving newlines.
  //
  // HTML in this repo should not rely on comment blocks to define required markup,
  // and guardrail tests should treat commented-out markup as non-existent.
  const text = String(html);
  let out = "";

  for (let i = 0; i < text.length; i += 1) {
    if (text.startsWith("<!--", i)) {
      out += "    "; // `<!--`
      i += 4;
      while (i < text.length && !text.startsWith("-->", i)) {
        out += text[i] === "\n" ? "\n" : " ";
        i += 1;
      }
      if (i < text.length) {
        out += "   "; // `-->`
        i += 2;
      } else {
        // Unterminated comment; stop scanning.
        break;
      }
      continue;
    }

    out += text[i];
  }

  return out;
}

export function stripHashComments(source) {
  // Strip `# ...` comments (YAML/shell-style) while preserving quoted strings and newlines.
  //
  // This is intentionally lightweight: it does not attempt to fully parse YAML or shell,
  // but is sufficient for guardrail tests that scan workflow files / scripts and should not
  // treat commented-out commands as present.
  const text = String(source);
  let out = "";
  /** @type {"code" | "single" | "double"} */
  let state = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : "";
    const prev = i > 0 ? text[i - 1] : "";

    if (state === "code") {
      if (ch === "'") {
        state = "single";
        out += ch;
        continue;
      }

      if (ch === '"') {
        state = "double";
        out += ch;
        continue;
      }

      const startsLine = i === 0 || prev === "\n";
      if (ch === "#" && (startsLine || /\s/.test(prev))) {
        // Treat as a comment until the end of the line.
        while (i < text.length && text[i] !== "\n") {
          out += " ";
          i += 1;
        }
        if (i < text.length) out += "\n";
        continue;
      }

      out += ch;
      continue;
    }

    // String literals: preserve as-is so `#` inside quotes doesn't get stripped.
    out += ch;
    if (state === "double" && ch === "\\") {
      if (next) {
        out += next;
        i += 1;
      }
      continue;
    }

    if (state === "single" && ch === "'" && next === "'") {
      // YAML-style escaping for single quotes: `''` => literal `'`.
      out += next;
      i += 1;
      continue;
    }

    if (state === "single" && ch === "'") {
      state = "code";
    } else if (state === "double" && ch === '"') {
      state = "code";
    }
  }

  return out;
}

export function stripPowerShellComments(source) {
  // Strip PowerShell line comments (`# ...`) and block comments (`<# ... #>`) while preserving:
  // - string literals (single/double quoted)
  // - here-strings (@' ... '@ / @" ... "@)
  // - newlines (so we don't accidentally join tokens)
  //
  // This is intentionally lightweight (not a full PowerShell parser), but is sufficient for
  // "source scanning" guardrail tests so commented-out script logic can't satisfy or fail checks.
  const text = String(source);
  let out = "";
  /** @type {"code" | "single" | "double" | "hereSingle" | "hereDouble" | "lineComment" | "blockComment"} */
  let state = "code";
  let lineStart = true;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : "";

    if (state === "lineComment") {
      if (ch === "\n") {
        out += "\n";
        state = "code";
        lineStart = true;
      } else {
        out += " ";
        lineStart = false;
      }
      continue;
    }

    if (state === "blockComment") {
      if (ch === "#" && next === ">") {
        out += "  ";
        i += 1;
        state = "code";
        lineStart = false;
        continue;
      }
      out += ch === "\n" ? "\n" : " ";
      lineStart = ch === "\n";
      continue;
    }

    if (state === "hereSingle" || state === "hereDouble") {
      const quote = state === "hereSingle" ? "'" : '"';
      if (lineStart && ch === quote && next === "@") {
        out += `${ch}${next}`;
        i += 1;
        state = "code";
        lineStart = false;
        continue;
      }
      out += ch;
      lineStart = ch === "\n";
      continue;
    }

    if (state === "single") {
      out += ch;
      if (ch === "'" && next === "'") {
        // PowerShell escapes a single quote by doubling it: `''` => literal `'`.
        out += next;
        i += 1;
        lineStart = false;
        continue;
      }
      if (ch === "'") state = "code";
      lineStart = ch === "\n";
      continue;
    }

    if (state === "double") {
      out += ch;
      if (ch === "`" && next) {
        // Backtick escapes the following character.
        out += next;
        i += 1;
        lineStart = next === "\n";
        continue;
      }
      if (ch === '"') state = "code";
      lineStart = ch === "\n";
      continue;
    }

    // state === "code"
    if (ch === "@" && (next === "'" || next === '"')) {
      // Here-string start: @' or @" (terminator must be `'@`/`"@` at start of line).
      out += `${ch}${next}`;
      i += 1;
      state = next === "'" ? "hereSingle" : "hereDouble";
      lineStart = false;
      continue;
    }

    if (ch === "<" && next === "#") {
      out += "  ";
      i += 1;
      state = "blockComment";
      lineStart = false;
      continue;
    }

    if (ch === "#") {
      out += " ";
      state = "lineComment";
      lineStart = false;
      continue;
    }

    if (ch === "'") {
      state = "single";
      out += ch;
      lineStart = false;
      continue;
    }

    if (ch === '"') {
      state = "double";
      out += ch;
      lineStart = false;
      continue;
    }

    out += ch;
    lineStart = ch === "\n";
  }

  return out;
}

export function stripPythonComments(source) {
  // Strip Python `# ...` line comments while preserving string literals (including triple quotes)
  // and newlines.
  //
  // This is intentionally lightweight: it does not attempt to parse all Python syntax, but is
  // sufficient for guardrail tests scanning `.py` sources where commented-out code must not
  // satisfy or fail assertions.
  const text = String(source);
  let out = "";
  /** @type {"code" | "single" | "double" | "tripleSingle" | "tripleDouble"} */
  let state = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : "";

    if (state === "code") {
      if (ch === "#") {
        // Line comment: replace with spaces until newline, then preserve the newline.
        while (i < text.length && text[i] !== "\n") {
          out += " ";
          i += 1;
        }
        if (i < text.length) out += "\n";
        continue;
      }

      if (ch === "'" || ch === '"') {
        const quote = ch;
        const isTriple = text[i + 1] === quote && text[i + 2] === quote;
        if (isTriple) {
          out += `${quote}${quote}${quote}`;
          i += 2;
          state = quote === "'" ? "tripleSingle" : "tripleDouble";
          continue;
        }
        out += quote;
        state = quote === "'" ? "single" : "double";
        continue;
      }

      out += ch;
      continue;
    }

    if (state === "single" || state === "double") {
      const quote = state === "single" ? "'" : '"';
      out += ch;
      if (ch === "\\") {
        // Escape sequence.
        if (next) {
          out += next;
          i += 1;
        }
        continue;
      }
      if (ch === quote) state = "code";
      continue;
    }

    if (state === "tripleSingle" || state === "tripleDouble") {
      const quote = state === "tripleSingle" ? "'" : '"';
      if (ch === quote && text[i + 1] === quote && text[i + 2] === quote) {
        out += `${quote}${quote}${quote}`;
        i += 2;
        state = "code";
        continue;
      }
      out += ch;
      continue;
    }
  }

  return out;
}
