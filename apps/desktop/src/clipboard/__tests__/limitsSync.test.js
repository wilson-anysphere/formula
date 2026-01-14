import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { CLIPBOARD_LIMITS } from "../platform/provider.js";
import { stripRustComments } from "../../../test/sourceTextUtils.js";

/**
 * Parse Rust `const`/`pub const` usize definitions from the clipboard module so we can keep JS
 * clipboard IPC guardrails in sync with the Rust backend.
 *
 * We intentionally keep this lightweight (no YAML/Rust parser dependencies) and only support
 * the expression subset used by the constants we care about (numbers, `+`, `*`, parentheses,
 * and identifiers referencing other constants).
 *
 * @param {string} text
 */
function parseUsizeConsts(text) {
  /** @type {Map<string, string>} */
  const out = new Map();
  const re = /^\s*(?:pub\s+)?const\s+([A-Z0-9_]+)\s*:\s*usize\s*=\s*([^;]+);/gm;
  for (const match of text.matchAll(re)) {
    const name = match[1];
    const expr = match[2]?.trim() ?? "";
    if (name && expr) out.set(name, expr);
  }
  return out;
}

/**
 * @typedef {{ type: "num", value: number } | { type: "ident", value: string } | { type: "op", value: string }} Token
 */

/**
 * @param {string} expr
 * @returns {Token[]}
 */
function tokenize(expr) {
  /** @type {Token[]} */
  const tokens = [];
  let i = 0;
  while (i < expr.length) {
    const ch = expr[i];
    if (!ch) break;
    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }
    if (ch === "+" || ch === "*" || ch === "(" || ch === ")") {
      tokens.push({ type: "op", value: ch });
      i += 1;
      continue;
    }
    if (/[0-9]/.test(ch)) {
      let j = i + 1;
      while (j < expr.length && /[0-9_]/.test(expr[j] ?? "")) j += 1;
      const raw = expr.slice(i, j).replace(/_/g, "");
      tokens.push({ type: "num", value: Number.parseInt(raw, 10) });
      i = j;
      continue;
    }
    if (/[A-Za-z_]/.test(ch)) {
      let j = i + 1;
      while (j < expr.length && /[A-Za-z0-9_]/.test(expr[j] ?? "")) j += 1;
      tokens.push({ type: "ident", value: expr.slice(i, j) });
      i = j;
      continue;
    }
    throw new Error(`Unexpected character in Rust const expr: ${JSON.stringify(ch)} in ${JSON.stringify(expr)}`);
  }
  return tokens;
}

/**
 * @param {Token[]} tokens
 * @param {Map<string, string>} consts
 * @param {Set<string>} stack
 */
function parseAndEval(tokens, consts, stack) {
  let pos = 0;

  /** @returns {Token | undefined} */
  function peek() {
    return tokens[pos];
  }

  /** @returns {Token} */
  function consume() {
    const t = tokens[pos];
    if (!t) {
      throw new Error("Unexpected end of Rust const expr.");
    }
    pos += 1;
    return t;
  }

  /** @returns {number} */
  function parseExpression() {
    let value = parseTerm();
    while (peek()?.type === "op" && peek()?.value === "+") {
      consume();
      value += parseTerm();
    }
    return value;
  }

  /** @returns {number} */
  function parseTerm() {
    let value = parseFactor();
    while (peek()?.type === "op" && peek()?.value === "*") {
      consume();
      value *= parseFactor();
    }
    return value;
  }

  /** @returns {number} */
  function parseFactor() {
    const t = consume();
    if (t.type === "num") return t.value;
    if (t.type === "ident") {
      // Reject function calls (identifier followed by '(').
      if (peek()?.type === "op" && peek()?.value === "(") {
        throw new Error(`Unsupported Rust const expr (function call): ${t.value}(...)`);
      }
      return evalConst(t.value, consts, stack);
    }
    if (t.type === "op" && t.value === "(") {
      const inner = parseExpression();
      const close = consume();
      if (close.type !== "op" || close.value !== ")") {
        throw new Error(`Expected ')' in Rust const expr, got ${close.type === "op" ? close.value : close.type}`);
      }
      return inner;
    }
    throw new Error(`Unexpected token in Rust const expr: ${JSON.stringify(t)}`);
  }

  const result = parseExpression();
  if (pos !== tokens.length) {
    throw new Error(`Unexpected trailing tokens in Rust const expr: ${tokens.slice(pos).map((t) => t.type).join(" ")}`);
  }
  return result;
}

/**
 * @param {string} name
 * @param {Map<string, string>} consts
 * @param {Set<string>} stack
 * @returns {number}
 */
function evalConst(name, consts, stack) {
  if (stack.has(name)) {
    throw new Error(`Cycle detected while evaluating Rust const ${name}`);
  }
  const expr = consts.get(name);
  if (!expr) throw new Error(`Missing Rust const ${name}`);
  stack.add(name);
  try {
    return parseAndEval(tokenize(expr), consts, stack);
  } finally {
    stack.delete(name);
  }
}

test("clipboard provider JS limits match Rust backend clipboard IPC guardrails", async () => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const repoRoot = path.resolve(here, "../../../../..");
  const rustPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "src", "clipboard", "mod.rs");
  const rustText = stripRustComments(await readFile(rustPath, "utf8"));

  const consts = parseUsizeConsts(rustText);
  const rustMaxPngBytes = evalConst("MAX_PNG_BYTES", consts, new Set());
  const rustMaxTextBytes = evalConst("MAX_TEXT_BYTES", consts, new Set());
  const rustMaxPlainTextWriteBytes = evalConst("MAX_PLAINTEXT_WRITE_BYTES", consts, new Set());

  assert.equal(CLIPBOARD_LIMITS.maxImageBytes, rustMaxPngBytes);
  assert.equal(CLIPBOARD_LIMITS.maxRichTextBytes, rustMaxTextBytes);
  assert.equal(CLIPBOARD_LIMITS.maxPlainTextWriteBytes, rustMaxPlainTextWriteBytes);
});
