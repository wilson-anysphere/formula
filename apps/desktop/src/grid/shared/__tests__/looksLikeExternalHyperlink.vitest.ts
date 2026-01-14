import { describe, expect, it } from "vitest";
import { looksLikeExternalHyperlink } from "../looksLikeExternalHyperlink";

describe("looksLikeExternalHyperlink", () => {
  it("accepts http(s) + mailto schemes", () => {
    expect(looksLikeExternalHyperlink("https://example.com")).toBe(true);
    expect(looksLikeExternalHyperlink("http://example.com")).toBe(true);
    expect(looksLikeExternalHyperlink("mailto:test@example.com")).toBe(true);
  });

  it("rejects non-URLs and empty strings", () => {
    expect(looksLikeExternalHyperlink("")).toBe(false);
    expect(looksLikeExternalHyperlink("   ")).toBe(false);
    expect(looksLikeExternalHyperlink("foo:bar")).toBe(false);
    expect(looksLikeExternalHyperlink("example.com")).toBe(false);
  });
});

