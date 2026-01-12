/**
 * Very small Excel-ish formula tokenizer.
 *
 * This intentionally supports only the subset needed for semantic diff:
 * - numbers, strings, identifiers, cell refs, basic operators, ranges, and function calls
 *
 * It is NOT intended to be a full Excel parser.
 */

/**
 * @typedef {"number"|"string"|"ident"|"op"|"punct"|"eof"} TokenType
 * @typedef {{ type: TokenType, value: string }} Token
 */

/**
 * @param {string} input
 * @returns {Token[]}
 */
export function tokenizeFormula(input) {
  /** @type {Token[]} */
  const tokens = [];
  let i = 0;

  const punct = new Set([",", "(", ")", ";", "{", "}", "[", "]"]);
  const singleCharOps = new Set([
    "+",
    "-",
    "*",
    "/",
    "^",
    ":",
    "!",
    "=",
    "%",
    "<",
    ">",
    "&",
    "@",
    "#",
  ]);

  const push = (type, value) => {
    tokens.push({ type, value });
  };

  while (i < input.length) {
    const ch = input[i];

    // whitespace
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      i += 1;
      continue;
    }

    // string literal: "foo" (Excel uses "" to escape ")
    if (ch === "\"") {
      let j = i + 1;
      let out = "";
      while (j < input.length) {
        const c = input[j];
        if (c === "\"") {
          // escaped quote
          if (input[j + 1] === "\"") {
            out += "\"";
            j += 2;
            continue;
          }
          break;
        }
        out += c;
        j += 1;
      }
      if (j >= input.length || input[j] !== "\"") {
        throw new Error("Unterminated string literal in formula");
      }
      push("string", out);
      i = j + 1;
      continue;
    }

    // quoted sheet name: 'My Sheet'!A1
    if (ch === "'") {
      let j = i + 1;
      let out = "";
      while (j < input.length) {
        const c = input[j];
        if (c === "'") {
          // escaped single quote is ''
          if (input[j + 1] === "'") {
            out += "'";
            j += 2;
            continue;
          }
          break;
        }
        out += c;
        j += 1;
      }
      if (j >= input.length || input[j] !== "'") {
        throw new Error("Unterminated quoted sheet name in formula");
      }
      // treat as identifier token so the parser can interpret sheet names
      push("ident", out);
      i = j + 1;
      continue;
    }

    // number: 12, 12.34, .5, 1e3, 1.2E-3
    if (
      (ch >= "0" && ch <= "9") ||
      (ch === "." && i + 1 < input.length && input[i + 1] >= "0" && input[i + 1] <= "9")
    ) {
      let j = i;
      let sawDot = false;

      if (input[j] === ".") {
        sawDot = true;
        j += 1;
      }

      while (j < input.length && input[j] >= "0" && input[j] <= "9") j += 1;

      if (!sawDot && input[j] === ".") {
        sawDot = true;
        j += 1;
        while (j < input.length && input[j] >= "0" && input[j] <= "9") j += 1;
      }

      // exponent
      if (input[j] === "e" || input[j] === "E") {
        let k = j + 1;
        if (input[k] === "+" || input[k] === "-") k += 1;
        const expStart = k;
        while (k < input.length && input[k] >= "0" && input[k] <= "9") k += 1;
        if (k !== expStart) {
          j = k;
        }
      }

      push("number", input.slice(i, j));
      i = j;
      continue;
    }

    // identifiers: letters/_/./$ + digits/_/./$
    if (
      (ch >= "A" && ch <= "Z") ||
      (ch >= "a" && ch <= "z") ||
      ch === "_" ||
      ch === "." ||
      ch === "$"
    ) {
      let j = i + 1;
      while (j < input.length) {
        const c = input[j];
        if (
          (c >= "A" && c <= "Z") ||
          (c >= "a" && c <= "z") ||
          (c >= "0" && c <= "9") ||
          c === "_" ||
          c === "." ||
          c === "$"
        ) {
          j += 1;
        } else {
          break;
        }
      }
      push("ident", input.slice(i, j));
      i = j;
      continue;
    }

    // operators / punctuation
    // Common multi-char operators
    if (ch === "<") {
      const nextCh = input[i + 1];
      if (nextCh === "=" || nextCh === ">") {
        push("op", `${ch}${nextCh}`);
        i += 2;
        continue;
      }
    }
    if (ch === ">") {
      const nextCh = input[i + 1];
      if (nextCh === "=") {
        push("op", `${ch}${nextCh}`);
        i += 2;
        continue;
      }
    }

    if (punct.has(ch)) {
      push("punct", ch);
      i += 1;
      continue;
    }
    if (singleCharOps.has(ch)) {
      // Treat % as a postfix operator.
      push("op", ch);
      i += 1;
      continue;
    }

    // Be permissive: keep unknown characters as operators so callers (diff UI)
    // can still render something, and higher-level parsers can decide whether
    // they understand the token stream.
    push("op", ch);
    i += 1;
  }

  push("eof", "");
  return tokens;
}
