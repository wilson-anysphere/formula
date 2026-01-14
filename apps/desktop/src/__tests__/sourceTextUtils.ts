export function skipStringLiteral(source: string, start: number): number {
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

function skipTemplateLiteral(source: string, start: number): number {
  // Template literals can contain nested `${ ... }` expressions (including other template literals).
  // This is a lightweight best-effort skipper so our source-scanning guardrails don't get confused
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
        const c = source[i]!;

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

function isRegexLiteralStart(source: string, start: number): boolean {
  // Best-effort detection of regex literals vs division.
  // This is intentionally heuristic (not a full JS lexer), but handles the patterns we
  // use in desktop guardrail tests (e.g. `.split(/[/\\\\]/)`).
  if (source[start] !== "/") return false;
  const next = source[start + 1];
  if (next == null) return false;
  // `//` and `/*` are always comments (regex bodies cannot start with `/` or `*`).
  if (next === "/" || next === "*") return false;

  let i = start - 1;
  while (i >= 0 && /\s/.test(source[i]!)) i -= 1;
  if (i < 0) return true;
  const prev = source[i]!;

  // Characters that can precede an expression, where a regex literal is valid.
  if ("([{:;,=!?&|+-*%^~<>/".includes(prev)) return true;

  // Keywords like `return /.../` are also valid.
  if (/[A-Za-z_$]/.test(prev)) {
    let j = i;
    while (j >= 0 && /[A-Za-z0-9_$]/.test(source[j]!)) j -= 1;
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

function skipRegexLiteral(source: string, start: number): number {
  // Assumes `source[start] === '/'` and `isRegexLiteralStart(source, start) === true`.
  let inCharClass = false;
  let escaped = false;

  for (let i = start + 1; i < source.length; i += 1) {
    const ch = source[i]!;
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
      while (j < source.length && /[A-Za-z]/.test(source[j]!)) j += 1;
      return j;
    }
  }

  return source.length;
}

export function stripComments(source: string): string {
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

export function stripCssComments(css: string): string {
  // Strip CSS block comments while preserving string literals and newlines.
  //
  // This is used by source-scanning guardrail tests so commented-out selectors/declarations
  // cannot satisfy or fail assertions.
  const text = String(css);
  let out = "";
  type State = "code" | "single" | "double" | "comment";
  let state: State = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]!;
    const next = i + 1 < text.length ? text[i + 1]! : "";

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

export function stripHtmlComments(html: string): string {
  // Strip HTML/XML comments (`<!-- ... -->`) while preserving newlines.
  //
  // Guardrail tests should not treat commented-out markup as present.
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
        break;
      }
      continue;
    }

    out += text[i]!;
  }

  return out;
}

export function stripHashComments(source: string): string {
  // Strip `# ...` comments (YAML/shell-style) while preserving quoted strings and newlines.
  //
  // This is a lightweight helper for workflow/docs guardrails that scan text files where
  // commented-out commands/flags must not satisfy assertions.
  const text = String(source);
  let out = "";
  type State = "code" | "single" | "double";
  let state: State = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]!;
    const next = i + 1 < text.length ? text[i + 1]! : "";
    const prev = i > 0 ? text[i - 1]! : "";

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
      // YAML-style single-quote escaping: `''` => literal `'`.
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

function parseRustRawStringStart(source: string, start: number): { endQuote: number; hashCount: number } | null {
  // Rust raw strings:
  // - r"..." / r#"..."# / r##"..."## / ...
  // - br"..." / br#"..."# / ...
  let i = start;
  if (source[i] === "b" && source[i + 1] === "r") {
    i += 2;
  } else if (source[i] === "r") {
    i += 1;
  } else {
    return null;
  }

  let hashCount = 0;
  while (source[i] === "#") {
    hashCount += 1;
    i += 1;
  }

  if (source[i] !== '"') return null;
  return { endQuote: i, hashCount };
}

export function stripRustComments(source: string): string {
  // Strip Rust `//` and `/* */` comments while preserving string literals and newlines.
  //
  // Rust block comments are nestable; this is a best-effort implementation sufficient for
  // source-scanning guardrails in tests (e.g. extracting event names / const definitions).
  const text = String(source);
  let out = "";
  type State = "code" | "string" | "rawString" | "lineComment" | "blockComment";
  let state: State = "code";
  let rawHashCount = 0;
  let blockDepth = 0;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]!;
    const next = i + 1 < text.length ? text[i + 1]! : "";

    if (state === "lineComment") {
      if (ch === "\n") {
        out += "\n";
        state = "code";
      } else {
        out += " ";
      }
      continue;
    }

    if (state === "blockComment") {
      if (ch === "/" && next === "*") {
        out += "  ";
        i += 1;
        blockDepth += 1;
        continue;
      }
      if (ch === "*" && next === "/") {
        out += "  ";
        i += 1;
        blockDepth -= 1;
        if (blockDepth <= 0) {
          blockDepth = 0;
          state = "code";
        }
        continue;
      }
      out += ch === "\n" ? "\n" : " ";
      continue;
    }

    if (state === "string") {
      out += ch;
      if (ch === "\\") {
        if (next) {
          out += next;
          i += 1;
        }
        continue;
      }
      if (ch === '"') {
        state = "code";
      }
      continue;
    }

    if (state === "rawString") {
      out += ch;
      if (ch === '"') {
        let ok = true;
        for (let j = 0; j < rawHashCount; j += 1) {
          if (text[i + 1 + j] !== "#") {
            ok = false;
            break;
          }
        }
        if (ok) {
          for (let j = 0; j < rawHashCount; j += 1) out += "#";
          i += rawHashCount;
          state = "code";
        }
      }
      continue;
    }

    // code
    const rawStart = parseRustRawStringStart(text, i);
    if (rawStart) {
      const { endQuote, hashCount } = rawStart;
      out += text.slice(i, endQuote + 1);
      i = endQuote;
      state = "rawString";
      rawHashCount = hashCount;
      continue;
    }

    if (ch === "b" && next === '"') {
      out += 'b"';
      i += 1;
      state = "string";
      continue;
    }

    if (ch === '"') {
      out += ch;
      state = "string";
      continue;
    }

    if (ch === "/" && next === "/") {
      out += "  ";
      i += 1;
      state = "lineComment";
      continue;
    }

    if (ch === "/" && next === "*") {
      out += "  ";
      i += 1;
      state = "blockComment";
      blockDepth = 1;
      continue;
    }

    out += ch;
  }

  return out;
}
