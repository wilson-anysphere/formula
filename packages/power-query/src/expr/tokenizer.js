import "./ast.js";

/**
 * @typedef {import("./ast.js").ExprToken} ExprToken
 */

/**
 * @param {string} message
 * @param {string} input
 * @param {number} pos
 * @returns {never}
 */
function syntaxError(message, input, pos) {
  const caret = `${" ".repeat(Math.max(0, pos))}^`;
  throw new Error(`${message}\n${input}\n${caret}`);
}

/**
 * @param {string} input
 * @param {number} pos
 * @returns {{ value: string, end: number }}
 */
function readColumnRef(input, pos) {
  const end = input.indexOf("]", pos + 1);
  if (end === -1) syntaxError("Unterminated column reference", input, pos);
  const raw = input.slice(pos + 1, end).trim();
  if (!raw) syntaxError("Empty column reference", input, pos);
  return { value: raw, end: end + 1 };
}

/**
 * @param {string} input
 * @param {number} pos
 * @returns {{ value: string, end: number }}
 */
function readStringLiteral(input, pos) {
  const quote = input[pos];
  let i = pos + 1;
  let out = "";

  for (; i < input.length; i++) {
    const ch = input[i];

    if (ch === quote) {
      return { value: out, end: i + 1 };
    }

    if (ch === "\\") {
      const esc = input[i + 1];
      if (esc == null) syntaxError("Unterminated escape sequence", input, i);
      i += 1;
      switch (esc) {
        case "\\":
        case '"':
        case "'":
          out += esc;
          break;
        case "n":
          out += "\n";
          break;
        case "r":
          out += "\r";
          break;
        case "t":
          out += "\t";
          break;
        case "b":
          out += "\b";
          break;
        case "f":
          out += "\f";
          break;
        case "v":
          out += "\v";
          break;
        case "0":
          out += "\0";
          break;
        case "u": {
          const hex = input.slice(i + 1, i + 5);
          if (!/^[0-9a-fA-F]{4}$/.test(hex)) {
            syntaxError("Invalid \\u escape sequence", input, i - 1);
          }
          out += String.fromCharCode(Number.parseInt(hex, 16));
          i += 4;
          break;
        }
        case "x": {
          const hex = input.slice(i + 1, i + 3);
          if (!/^[0-9a-fA-F]{2}$/.test(hex)) {
            syntaxError("Invalid \\x escape sequence", input, i - 1);
          }
          out += String.fromCharCode(Number.parseInt(hex, 16));
          i += 2;
          break;
        }
        default:
          syntaxError(`Unsupported escape sequence \\${esc}`, input, i - 1);
      }
      continue;
    }

    if (ch === "\n" || ch === "\r") {
      syntaxError("Unterminated string literal", input, pos);
    }

    out += ch;
  }

  syntaxError("Unterminated string literal", input, pos);
}

/**
 * @param {string} input
 * @param {number} pos
 * @returns {{ value: number, end: number }}
 */
function readNumberLiteral(input, pos) {
  let i = pos;
  const start = pos;

  /** @param {string} ch */
  const isDigit = (ch) => ch >= "0" && ch <= "9";

  if (input[i] === ".") {
    i += 1;
    while (isDigit(input[i] ?? "")) i += 1;
  } else {
    while (isDigit(input[i] ?? "")) i += 1;
    if (input[i] === ".") {
      i += 1;
      while (isDigit(input[i] ?? "")) i += 1;
    }
  }

  const exp = input[i];
  if (exp === "e" || exp === "E") {
    const sign = input[i + 1];
    let j = i + 1;
    if (sign === "+" || sign === "-") j += 1;
    if (!isDigit(input[j] ?? "")) {
      syntaxError("Invalid exponent in number literal", input, i);
    }
    j += 1;
    while (isDigit(input[j] ?? "")) j += 1;
    i = j;
  }

  const raw = input.slice(start, i);
  const num = Number(raw);
  if (!Number.isFinite(num)) {
    syntaxError("Number literal must be finite", input, start);
  }

  return { value: num, end: i };
}

/**
 * @param {string} input
 * @param {number} pos
 * @returns {{ value: string, end: number }}
 */
function readIdentifier(input, pos) {
  let i = pos;
  while (i < input.length && /[A-Za-z0-9_]/.test(input[i])) i += 1;
  return { value: input.slice(pos, i), end: i };
}

const MULTI_CHAR_OPERATORS = ["!==", "===", "!=", "==", "<=", ">=", "||", "&&"];
const SINGLE_CHAR_OPERATORS = new Set(["+", "-", "*", "/", "%", "(", ")", "<", ">", "!", "?", ":", ","]);

/**
 * Tokenize an expression string.
 *
 * @param {string} input
 * @returns {ExprToken[]}
 */
export function tokenizeExpr(input) {
  /** @type {ExprToken[]} */
  const tokens = [];

  let i = 0;
  while (i < input.length) {
    const ch = input[i];

    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    if (ch === "[") {
      const start = i;
      const { value, end } = readColumnRef(input, i);
      tokens.push({ type: "column", value, span: { start, end } });
      i = end;
      continue;
    }

    if (ch === "'" || ch === '"') {
      const start = i;
      const { value, end } = readStringLiteral(input, i);
      tokens.push({ type: "string", value, span: { start, end } });
      i = end;
      continue;
    }

    if ((ch >= "0" && ch <= "9") || (ch === "." && /[0-9]/.test(input[i + 1] ?? ""))) {
      const start = i;
      const { value, end } = readNumberLiteral(input, i);
      tokens.push({ type: "number", value, span: { start, end } });
      i = end;
      continue;
    }

    if (/[A-Za-z_]/.test(ch)) {
      const start = i;
      const { value, end } = readIdentifier(input, i);
      tokens.push({ type: "identifier", value, span: { start, end } });
      i = end;
      continue;
    }

    const rest = input.slice(i);
    const multi = MULTI_CHAR_OPERATORS.find((op) => rest.startsWith(op));
    if (multi) {
      const start = i;
      const end = i + multi.length;
      tokens.push({ type: "operator", value: multi, span: { start, end } });
      i = end;
      continue;
    }

    if (SINGLE_CHAR_OPERATORS.has(ch)) {
      const start = i;
      const end = i + 1;
      tokens.push({ type: "operator", value: ch, span: { start, end } });
      i = end;
      continue;
    }

    syntaxError(`Unsupported character '${ch}'`, input, i);
  }

  tokens.push({ type: "eof", span: { start: input.length, end: input.length } });
  return tokens;
}

