import { describe, expect, it } from "vitest";

import { stripComments } from "./sourceTextUtils";

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
});

