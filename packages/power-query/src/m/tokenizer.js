import { MLanguageSyntaxError } from "./errors.js";

/**
 * @typedef {"keyword" | "identifier" | "quotedIdentifier" | "string" | "number" | "operator" | "punct" | "eof"} TokenType
 */

/**
 * @typedef {{
 *   type: TokenType;
 *   value: string;
 *   start: import("./errors.js").MLocation;
 *   end: import("./errors.js").MLocation;
 * }} Token
 */

const KEYWORDS = new Set([
  "let",
  "in",
  "each",
  "and",
  "or",
  "not",
  "true",
  "false",
  "null",
  "type",
  "if",
  "then",
  "else",
  "try",
  "otherwise",
  "as",
]);

/**
 * @param {string} char
 * @returns {boolean}
 */
function isIdentifierStart(char) {
  return /[A-Za-z_]/.test(char);
}

/**
 * @param {string} char
 * @returns {boolean}
 */
function isIdentifierContinue(char) {
  return /[A-Za-z0-9_]/.test(char);
}

/**
 * @param {string} source
 * @returns {Token[]}
 */
export function tokenizeM(source) {
  /** @type {Token[]} */
  const tokens = [];

  let offset = 0;
  let line = 1;
  let column = 1;

  /**
   * @returns {import("./errors.js").MLocation}
   */
  const loc = () => ({ offset, line, column });

  /**
   * @param {number} n
   */
  function advance(n = 1) {
    for (let i = 0; i < n; i++) {
      const char = source[offset];
      offset += 1;
      if (char === "\n") {
        line += 1;
        column = 1;
      } else {
        column += 1;
      }
    }
  }

  /**
   * @param {TokenType} type
   * @param {string} value
   * @param {import("./errors.js").MLocation} start
   * @param {import("./errors.js").MLocation} end
   */
  function push(type, value, start, end) {
    tokens.push({ type, value, start, end });
  }

  /**
   * @param {string} message
   * @param {import("./errors.js").MLocation} at
   */
  function error(message, at) {
    throw new MLanguageSyntaxError(message, { location: at, source });
  }

  /**
   * @returns {boolean}
   */
  function eof() {
    return offset >= source.length;
  }

  while (!eof()) {
    const char = source[offset];

    // Whitespace
    if (char === " " || char === "\t" || char === "\n" || char === "\r") {
      advance(1);
      continue;
    }

    // Line comment
    if (char === "/" && source[offset + 1] === "/") {
      while (!eof() && source[offset] !== "\n") advance(1);
      continue;
    }

    // Block comment
    if (char === "/" && source[offset + 1] === "*") {
      advance(2);
      while (!eof() && !(source[offset] === "*" && source[offset + 1] === "/")) {
        advance(1);
      }
      if (eof()) error("Unterminated block comment", loc());
      advance(2);
      continue;
    }

    const start = loc();

    // Quoted identifier: #"..."
    if (char === "#" && source[offset + 1] === '"') {
      advance(2);
      let value = "";
      while (!eof()) {
        const c = source[offset];
        if (c === '"') {
          const next = source[offset + 1];
          if (next === '"') {
            value += '"';
            advance(2);
            continue;
          }
          advance(1);
          push("quotedIdentifier", value, start, loc());
          value = "";
          break;
        }
        value += c;
        advance(1);
      }
      if (value !== "") error("Unterminated quoted identifier", start);
      continue;
    }

    // String: "..."
    if (char === '"') {
      advance(1);
      let value = "";
      while (!eof()) {
        const c = source[offset];
        if (c === '"') {
          const next = source[offset + 1];
          if (next === '"') {
            value += '"';
            advance(2);
            continue;
          }
          advance(1);
          push("string", value, start, loc());
          value = "";
          break;
        }
        value += c;
        advance(1);
      }
      if (value !== "") error("Unterminated string literal", start);
      continue;
    }

    // String: '...' (best-effort; not part of canonical M syntax but appears in some generators)
    if (char === "'") {
      advance(1);
      let value = "";
      while (!eof()) {
        const c = source[offset];
        if (c === "'") {
          const next = source[offset + 1];
          if (next === "'") {
            value += "'";
            advance(2);
            continue;
          }
          advance(1);
          push("string", value, start, loc());
          value = "";
          break;
        }
        value += c;
        advance(1);
      }
      if (value !== "") error("Unterminated string literal", start);
      continue;
    }

    // Hash identifiers (e.g. #date)
    if (char === "#" && isIdentifierStart(source[offset + 1] ?? "")) {
      advance(1);
      let name = "#";
      while (!eof() && isIdentifierContinue(source[offset])) {
        name += source[offset];
        advance(1);
      }
      push("identifier", name, start, loc());
      continue;
    }

    // Number
    if (/[0-9]/.test(char)) {
      let raw = "";
      while (!eof() && /[0-9]/.test(source[offset])) {
        raw += source[offset];
        advance(1);
      }
      if (source[offset] === "." && /[0-9]/.test(source[offset + 1] ?? "")) {
        raw += ".";
        advance(1);
        while (!eof() && /[0-9]/.test(source[offset])) {
          raw += source[offset];
          advance(1);
        }
      }
      // Exponent (best-effort)
      if (/e/i.test(source[offset] ?? "")) {
        const expStart = offset;
        let exp = source[offset];
        advance(1);
        if (source[offset] === "+" || source[offset] === "-") {
          exp += source[offset];
          advance(1);
        }
        if (!/[0-9]/.test(source[offset] ?? "")) {
          // Roll back; it wasn't really an exponent.
          offset = expStart;
          column -= exp.length;
        } else {
          while (!eof() && /[0-9]/.test(source[offset])) {
            exp += source[offset];
            advance(1);
          }
          raw += exp;
        }
      }
      push("number", raw, start, loc());
      continue;
    }

    // Identifier / keyword
    if (isIdentifierStart(char)) {
      let ident = "";
      while (!eof() && isIdentifierContinue(source[offset])) {
        ident += source[offset];
        advance(1);
      }
      if (KEYWORDS.has(ident)) {
        push("keyword", ident, start, loc());
      } else {
        push("identifier", ident, start, loc());
      }
      continue;
    }

    // Operators / punctuation
    const twoChar = source.slice(offset, offset + 2);
    if (twoChar === "<>" || twoChar === "<=" || twoChar === ">=" || twoChar === "=>") {
      advance(2);
      push("operator", twoChar, start, loc());
      continue;
    }

    // Single-character tokens.
    const singleOperators = new Set(["=", "<", ">", "+", "-", "*", "/", "&"]);
    const punct = new Set(["(", ")", "{", "}", "[", "]", ",", "."]);

    if (singleOperators.has(char)) {
      advance(1);
      push("operator", char, start, loc());
      continue;
    }

    if (punct.has(char)) {
      advance(1);
      push("punct", char, start, loc());
      continue;
    }

    error(`Unexpected character '${char}'`, start);
  }

  const end = loc();
  tokens.push({ type: "eof", value: "", start: end, end });
  return tokens;
}
