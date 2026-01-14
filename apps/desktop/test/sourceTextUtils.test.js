import assert from "node:assert/strict";
import test from "node:test";

import { stripComments, stripCssComments, stripHtmlComments } from "./sourceTextUtils.js";

test("stripComments strips line/block comments but preserves string literals", () => {
  const input = [
    `const url = "https://example.com/path"; // trailing comment`,
    `const a = 1; /* block`,
    `comment */ const b = 2;`,
  ].join("\n");

  const out = stripComments(input);
  assert.match(out, /"https:\/\/example\.com\/path"/);
  assert.doesNotMatch(out, /\btrailing comment\b/);
  assert.doesNotMatch(out, /\bblock\b/);
  assert.doesNotMatch(out, /\bcomment \*\//);
  assert.match(out, /\bconst b = 2;/);
});

test("stripComments does not treat escaped slashes in regex literals as comment markers", () => {
  const input = String.raw`const re = /foo\//; // should be stripped`;
  const out = stripComments(input);
  assert.match(out, /\/foo\\\//);
  assert.doesNotMatch(out, /\bshould be stripped\b/);
});

test("stripComments does not strip comment markers inside regex character classes", () => {
  const input = `const re1 = /[/*]/; const re2 = /[//]/;`;
  const out = stripComments(input);
  assert.match(out, /\/\[\/\*\]\//);
  assert.match(out, /\/\[\/\/\]\//);
});

test("stripComments handles nested template literals and template expressions", () => {
  const input = [
    // Nested template literal inside an expression.
    "const a = `outer ${`inner`} end`; // trailing",
    // Block comment inside the expression (including a `}` that should not close the expression).
    "const b = `x ${foo /* } */} y`;",
    // Nested object literal braces inside the expression.
    "const c = `x ${ { a: 1, b: 2 } } y`;",
  ].join("\n");

  const out = stripComments(input);
  assert.match(out, /const a = `outer \${`inner`} end`;/);
  assert.doesNotMatch(out, /\btrailing\b/);
  assert.match(out, /const b = `x \${foo \/\* \} \*\/} y`;/);
  assert.match(out, /const c = `x \${ \{ a: 1, b: 2 \} } y`;/);
});

test("stripComments recognizes regex literals preceded by division operators", () => {
  // This pattern (a / /re/.test(x)) can trick simplistic scanners into treating `//` as a line comment
  // after the `/` operator. Ensure we keep the regex intact and only strip the trailing comment.
  // Use an escaped trailing slash in the regex (`/re\\//`) so the source contains `\//` (a common edge-case
  // for naive `//` comment scanners).
  const input = "const ok = a / /re\\//.test(x); // comment";
  const out = stripComments(input);
  assert.ok(
    out.includes("a / /re\\//.test(x);"),
    `expected stripped source to preserve the regex literal; got:\n${out}`,
  );
  assert.doesNotMatch(out, /\bcomment\b/);
});

test("stripCssComments strips block comments but preserves strings", () => {
  const input = [
    `.a { content: "/* not a comment */"; }`,
    `/* commented-out selector should not count: .b { color: red; } */`,
    `.c { color: var(--text-primary); }`,
    `.d[data-theme="high-contrast"] { outline: 1px solid var(--border); }`,
  ].join("\n");

  const out = stripCssComments(input);
  assert.match(out, /\.a\s*\{/);
  assert.match(out, /"\/\* not a comment \*\/"/);
  assert.doesNotMatch(out, /\.b\s*\{/);
  assert.match(out, /\.c\s*\{/);
  assert.match(out, /\.d\[data-theme="high-contrast"\]/);
});

test("stripHtmlComments strips HTML comments", () => {
  const input = [
    `<div id="app"></div>`,
    `<!-- <div id="commented"></div> -->`,
    `<div data-testid="live"></div>`,
  ].join("\n");

  const out = stripHtmlComments(input);
  assert.match(out, /\bid="app"/);
  assert.doesNotMatch(out, /\bid="commented"/);
  assert.match(out, /data-testid="live"/);
});
