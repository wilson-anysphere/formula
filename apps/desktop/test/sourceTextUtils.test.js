import assert from "node:assert/strict";
import test from "node:test";

import {
  stripComments,
  stripCssComments,
  stripHtmlComments,
  stripHashComments,
  stripYamlBlockScalarBodies,
  stripPowerShellComments,
  stripPythonComments,
  stripRustComments,
} from "./sourceTextUtils.js";

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

test("stripComments strips full-line comments adjacent to regex literals (does not get confused by regex literal heuristics)", () => {
  // Repro for a tricky case where a regex literal is preceded by a full-line comment:
  //
  //   // comment
  //   /re/
  //
  // The regex literal start detection must ignore the comment line so it doesn't misclassify
  // the opening/closing `/` and accidentally preserve later `// ...` comments.
  const input = [
    "const xs = [",
    "  // comment A",
    "  /a\\(/,",
    "  // comment B",
    "  /b\\(/,",
    "];",
  ].join("\n");
  const out = stripComments(input);
  assert.doesNotMatch(out, /\bcomment A\b/);
  assert.doesNotMatch(out, /\bcomment B\b/);
  assert.match(out, /\/a\\\(\//);
  assert.match(out, /\/b\\\(\//);
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

test("stripHashComments strips # line comments but preserves strings/urls", () => {
  const input = [
    `run: pnpm -C apps/desktop check:coi # --no-build (commented-out flag should not count)`,
    `  # pnpm -C apps/desktop check:coi --no-build (commented-out command should not count)`,
    `name: "build #1"`,
    `url: https://example.com/#anchor`,
    `single: 'it''s # not a comment'`,
  ].join("\n");

  const out = stripHashComments(input);
  assert.ok(out.includes("pnpm -C apps/desktop check:coi"), "expected the COI command itself to remain");
  assert.doesNotMatch(out, /--no-build \(commented-out flag/, "expected inline # comment to be stripped");
  assert.doesNotMatch(out, /pnpm -C apps\/desktop check:coi --no-build \(commented-out command/, "expected full comment line to be stripped");
  assert.match(out, /name:\s*"build #1"/);
  assert.match(out, /https:\/\/example\.com\/#anchor/);
  assert.match(out, /single:\s*'it''s # not a comment'/);
});

test("stripYamlBlockScalarBodies strips YAML block scalar bodies", () => {
  const input = [
    "steps:",
    "  - name: First",
    "    run: |",
    "      echo \"name: Should not match\"",
    "      echo ok",
    "  - name: Second",
    "    run: echo hi",
  ].join("\n");

  const out = stripYamlBlockScalarBodies(stripHashComments(input));
  assert.match(out, /\brun:\s*\|/);
  assert.doesNotMatch(out, /\bShould not match\b/);
  assert.doesNotMatch(out, /\becho ok\b/);
  assert.match(out, /^\s*-\s*name:\s*Second\b/m);
  assert.match(out, /\brun: echo hi\b/);
});

test("stripYamlBlockScalarBodies does not end a block scalar early on blank lines", () => {
  const input = [
    "run: |",
    "  echo a",
    "",
    "  echo b",
    "name: After",
  ].join("\n");

  const out = stripYamlBlockScalarBodies(stripHashComments(input));
  assert.doesNotMatch(out, /\becho a\b/);
  assert.doesNotMatch(out, /\becho b\b/);
  assert.match(out, /\bname: After\b/);
});

test("stripPowerShellComments strips # and <# #> comments but preserves strings + here-strings", () => {
  const input = [
    `$x = $env:FORMULA_TAURI_CONF_PATH # trailing comment`,
    `# full line comment: $env:FORMULA_TAURI_CONF_PATH`,
    `<#`,
    `block comment: $env:FORMULA_TAURI_CONF_PATH`,
    `#>`,
    `Write-Host "FORMULA_TAURI_CONF_PATH # not a comment"`,
    `@'`,
    `here-string content # not a comment`,
    `'@`,
  ].join("\n");

  const out = stripPowerShellComments(input);
  assert.match(out, /\$env:FORMULA_TAURI_CONF_PATH/);
  assert.doesNotMatch(out, /\btrailing comment\b/);
  assert.doesNotMatch(out, /\bfull line comment\b/);
  assert.doesNotMatch(out, /\bblock comment\b/);
  assert.match(out, /"FORMULA_TAURI_CONF_PATH # not a comment"/);
  assert.match(out, /here-string content # not a comment/);
});

test("stripPythonComments strips # comments but preserves strings (including triple quotes)", () => {
  const input = [
    `x = 1  # trailing`,
    `y = "# not a comment"`,
    `z = """`,
    `triple string # not a comment`,
    `"""`,
    `# full line`,
  ].join("\n");

  const out = stripPythonComments(input);
  assert.doesNotMatch(out, /\btrailing\b/);
  assert.doesNotMatch(out, /\bfull line\b/);
  assert.match(out, /"# not a comment"/);
  assert.match(out, /triple string # not a comment/);
});

test("stripRustComments strips Rust line/block comments but preserves strings/raw strings", () => {
  const input = [
    `pub const MAX_BYTES: usize = 5 * 1024; // comment`,
    `/* block comment`,
    `pub const SHOULD_NOT_MATCH: usize = 1;`,
    `*/`,
    `pub const URL: &str = "https://example.com/#anchor";`,
    `pub const RAW: &str = r#"// not a comment"#;`,
  ].join("\n");

  const out = stripRustComments(input);
  assert.match(out, /\bpub const MAX_BYTES\b/);
  assert.doesNotMatch(out, /\/\/\s*comment\b/);
  assert.doesNotMatch(out, /\bblock comment\b/);
  assert.doesNotMatch(out, /\bSHOULD_NOT_MATCH\b/);
  assert.match(out, /https:\/\/example\.com\/#anchor/);
  assert.match(out, /r#\"\/\/ not a comment\"#/);
});
