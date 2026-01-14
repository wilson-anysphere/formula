import { describe, expect, it } from "vitest";

import { stripComments, stripCssComments, stripHashComments, stripHtmlComments, stripRustComments } from "./sourceTextUtils";

describe("sourceTextUtils.stripComments", () => {
  it("strips line/block comments but preserves string literals", () => {
    const input = [
      `const url = "https://example.com/path"; // trailing comment`,
      `const a = 1; /* block`,
      `comment */ const b = 2;`,
    ].join("\n");

    const out = stripComments(input);
    expect(out).toContain(`"https://example.com/path"`);
    expect(out).not.toContain("trailing comment");
    expect(out).not.toContain("block");
    expect(out).not.toContain("comment */");
    expect(out).toContain("const b = 2;");
  });

  it("does not treat escaped slashes in regex literals as comment markers", () => {
    const input = `const re = /foo\\//; // should be stripped`;
    const out = stripComments(input);
    expect(out).toContain(`/foo\\//`);
    expect(out).not.toContain("should be stripped");
  });

  it("does not strip comment markers inside regex character classes", () => {
    const input = `const re1 = /[/*]/; const re2 = /[//]/;`;
    const out = stripComments(input);
    expect(out).toContain(`/[/*]/`);
    expect(out).toContain(`/[//]/`);
  });

  it("handles nested template literals and template expressions without getting confused by backticks/braces", () => {
    const input = [
      // Nested template literal inside an expression.
      "const a = `outer ${`inner`} end`; // trailing",
      // Block comment inside the expression (including a `}` that should not close the expression).
      "const b = `x ${foo /* } */} y`;",
      // Nested object literal braces inside the expression.
      "const c = `x ${ { a: 1, b: 2 } } y`;",
    ].join("\n");

    const out = stripComments(input);
    expect(out).toContain("const a = `outer ${`inner`} end`;");
    expect(out).not.toContain("trailing");
    expect(out).toContain("const b = `x ${foo /* } */} y`;");
    expect(out).toContain("const c = `x ${ { a: 1, b: 2 } } y`;");
  });
});

describe("sourceTextUtils.stripCssComments", () => {
  it("strips block comments but preserves strings", () => {
    const input = [
      `.a { content: "/* not a comment */"; }`,
      `/* commented-out selector should not count: .b { color: red; } */`,
      `.c { color: var(--text-primary); }`,
    ].join("\n");

    const out = stripCssComments(input);
    expect(out).toContain(".a {");
    expect(out).toContain(`"/* not a comment */"`);
    expect(out).not.toContain(".b {");
    expect(out).toContain(".c {");
  });
});

describe("sourceTextUtils.stripHtmlComments", () => {
  it("strips <!-- ... --> comments", () => {
    const input = [`<div id="app"></div>`, `<!-- <div id="commented"></div> -->`, `<div data-testid="live"></div>`].join("\n");
    const out = stripHtmlComments(input);
    expect(out).toContain('id="app"');
    expect(out).not.toContain('id="commented"');
    expect(out).toContain('data-testid="live"');
  });
});

describe("sourceTextUtils.stripHashComments", () => {
  it("strips # line comments but preserves strings/urls", () => {
    const input = [
      `run: pnpm -C apps/desktop check:coi # --no-build (commented-out flag should not count)`,
      `  # pnpm -C apps/desktop check:coi --no-build (commented-out command should not count)`,
      `name: "build #1"`,
      `url: https://example.com/#anchor`,
      `single: 'it''s # not a comment'`,
    ].join("\n");

    const out = stripHashComments(input);
    expect(out).toContain("pnpm -C apps/desktop check:coi");
    expect(out).not.toContain("(commented-out flag should not count)");
    expect(out).not.toContain("(commented-out command should not count)");
    expect(out).toContain('name: "build #1"');
    expect(out).toContain("https://example.com/#anchor");
    expect(out).toContain("single: 'it''s # not a comment'");
  });
});

describe("sourceTextUtils.stripRustComments", () => {
  it("strips Rust line/block comments but preserves strings/raw strings", () => {
    const input = [
      `pub const MAX_BYTES: usize = 5 * 1024; // comment`,
      `/* block comment`,
      `pub const SHOULD_NOT_MATCH: usize = 1;`,
      `*/`,
      `pub const URL: &str = "https://example.com/#anchor";`,
      `pub const RAW: &str = r#"// not a comment"#;`,
    ].join("\n");

    const out = stripRustComments(input);
    expect(out).toContain("pub const MAX_BYTES");
    expect(out).not.toMatch(/\/\/\s*comment\b/);
    expect(out).not.toContain("block comment");
    expect(out).not.toContain("SHOULD_NOT_MATCH");
    expect(out).toContain("https://example.com/#anchor");
    expect(out).toContain(`r#"// not a comment"#`);
  });
});
