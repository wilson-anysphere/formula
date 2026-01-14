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

  it("returns localized function names in fr-FR (SOMME)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "fr-FR";

    try {
      const sections = buildCommandPaletteSections({
        query: "somme",
        commands: [],
        limits: { maxResults: 50, maxResultsPerGroup: 50 },
      });
      const functions = sections.find((s) => s.title === "FUNCTIONS");
      expect(functions).toBeTruthy();
      const names = functions!.results.filter((r) => r.kind === "function").map((r) => r.name);
      expect(names[0]).toBe("SOMME");
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("returns localized function names in es-ES (SUMA)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "es-ES";

    try {
      const sections = buildCommandPaletteSections({
        query: "suma",
        commands: [],
        limits: { maxResults: 50, maxResultsPerGroup: 50 },
      });
      const functions = sections.find((s) => s.title === "FUNCTIONS");
      expect(functions).toBeTruthy();
      const names = functions!.results.filter((r) => r.kind === "function").map((r) => r.name);
      expect(names[0]).toBe("SUMA");
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

  it("normalizes locale variants to the supported formula locale (de / de_DE.UTF-8 / de-CH-1996 → de-DE)", () => {
    const prevLang = document.documentElement.lang;

    try {
      for (const lang of ["de", "de_DE.UTF-8", "de-CH-1996"]) {
        document.documentElement.lang = lang;
        const results = searchFunctionResults("summe", { limit: 10 });
        expect(results[0]?.name).toBe("SUMME");
        // The engine treats these German variants as `de-DE` today, so list separators should match.
        expect(results[0]?.signature ?? "").toContain(";");
      }
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("treats unsupported locales as en-US for signature separators (pt-BR → comma)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "pt-BR";

    try {
      const results = searchFunctionResults("sum", { limit: 10 });
      expect(results[0]?.name).toBe("SUM");
      // The formula engine does not currently support pt-BR, so formula punctuation falls back to en-US.
      // That means `,` is used as the list/argument separator (even though pt-BR typically uses `;` in Excel).
      expect(results[0]?.signature ?? "").toContain(",");
      expect(results[0]?.signature ?? "").not.toContain(";");
    } finally {
      document.documentElement.lang = prevLang;
    }
  });
});
