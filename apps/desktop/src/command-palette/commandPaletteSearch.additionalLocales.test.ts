/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { buildCommandPaletteSections, searchFunctionResults } from "./commandPaletteSearch.js";

describe("command palette function search (localized function names)", () => {
  it("returns localized function names when document.lang is a supported formula locale (de-DE SUMME)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const sections = buildCommandPaletteSections({
        query: "summe",
        commands: [],
        limits: { maxResults: 50, maxResultsPerGroup: 50 },
      });
      const functions = sections.find((s) => s.title === "FUNCTIONS");
      expect(functions).toBeTruthy();
      const names = functions!.results.filter((r) => r.kind === "function").map((r) => r.name);
      expect(names[0]).toBe("SUMME");
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("supports non-ASCII queries for localized names (de-DE zähl → ZÄHLENWENN)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const sections = buildCommandPaletteSections({
        query: "zähl",
        commands: [],
        limits: { maxResults: 50, maxResultsPerGroup: 50 },
      });
      const functions = sections.find((s) => s.title === "FUNCTIONS");
      expect(functions).toBeTruthy();
      const names = functions!.results.filter((r) => r.kind === "function").map((r) => r.name);
      expect(names).toContain("ZÄHLENWENN");
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("respects the explicit localeId override (even when document.lang differs)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "en-US";

    try {
      const results = searchFunctionResults("zähl", { limit: 50, localeId: "de-DE" });
      expect(results.map((r) => r.name)).toContain("ZÄHLENWENN");
    } finally {
      document.documentElement.lang = prevLang;
    }
  });
});
