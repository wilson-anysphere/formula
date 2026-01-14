import { describe, expect, it } from "vitest";

import { normalizeFormulaLocaleId, normalizeLocaleId } from "./formulaLocale.js";

describe("formulaLocale", () => {
  it("normalizeLocaleId handles POSIX + BCP-47 locale ID variants", () => {
    expect(normalizeLocaleId("en-US")).toBe("en-US");
    expect(normalizeLocaleId("  de-DE  ")).toBe("de-DE");
    expect(normalizeLocaleId("de_DE.UTF-8")).toBe("de-DE");
    expect(normalizeLocaleId("de_DE@euro")).toBe("de-DE");
    expect(normalizeLocaleId("C")).toBe("en-US");
    expect(normalizeLocaleId("C.UTF-8")).toBe("en-US");
    expect(normalizeLocaleId("POSIX")).toBe("en-US");
    expect(normalizeLocaleId("fr-FR-u-nu-latn")).toBe("fr-FR");
    expect(normalizeLocaleId("de-CH-1996")).toBe("de-CH");
    expect(normalizeLocaleId("en")).toBe("en");
    expect(normalizeLocaleId("")).toBeNull();
    expect(normalizeLocaleId("   ")).toBeNull();
  });

  it("normalizeFormulaLocaleId maps language/variant IDs onto the engine's supported formula locales", () => {
    // Exact IDs still work.
    expect(normalizeFormulaLocaleId("en-US")).toBe("en-US");
    expect(normalizeFormulaLocaleId("de-DE")).toBe("de-DE");
    expect(normalizeFormulaLocaleId("fr-FR")).toBe("fr-FR");
    expect(normalizeFormulaLocaleId("es-ES")).toBe("es-ES");

    // Trim whitespace and treat `-` / `_` as equivalent (common OS spellings).
    expect(normalizeFormulaLocaleId("  de-DE  ")).toBe("de-DE");
    expect(normalizeFormulaLocaleId("de_DE")).toBe("de-DE");
    // Match case-insensitively.
    expect(normalizeFormulaLocaleId("DE-de")).toBe("de-DE");

    // POSIX locale IDs with encoding/modifier suffix.
    expect(normalizeFormulaLocaleId("de_DE.UTF-8")).toBe("de-DE");
    expect(normalizeFormulaLocaleId("de_DE@euro")).toBe("de-DE");

    // Language-only fallbacks.
    expect(normalizeFormulaLocaleId("de")).toBe("de-DE");
    expect(normalizeFormulaLocaleId("fr")).toBe("fr-FR");
    expect(normalizeFormulaLocaleId("es")).toBe("es-ES");
    expect(normalizeFormulaLocaleId("en")).toBe("en-US");

    // Language/region fallbacks.
    expect(normalizeFormulaLocaleId("fr-CA")).toBe("fr-FR");
    expect(normalizeFormulaLocaleId("de-AT")).toBe("de-DE");
    expect(normalizeFormulaLocaleId("es-MX")).toBe("es-ES");

    // The engine treats `en-GB` as an alias for the canonical formula locale.
    expect(normalizeFormulaLocaleId("en-GB")).toBe("en-US");
    expect(normalizeFormulaLocaleId("en-UK")).toBe("en-US");
    expect(normalizeFormulaLocaleId("en-AU")).toBe("en-US");
    expect(normalizeFormulaLocaleId("en-NZ")).toBe("en-US");

    // Ignore BCP-47 variants/extensions.
    expect(normalizeFormulaLocaleId("fr-FR-u-nu-latn")).toBe("fr-FR");
    expect(normalizeFormulaLocaleId("de-CH-1996")).toBe("de-DE");

    // Minimal locale registrations (no translations, but still valid engine locale ids).
    expect(normalizeFormulaLocaleId("ja")).toBe("ja-JP");
    expect(normalizeFormulaLocaleId("ja-JP")).toBe("ja-JP");
    expect(normalizeFormulaLocaleId("ko")).toBe("ko-KR");
    expect(normalizeFormulaLocaleId("ko-KR")).toBe("ko-KR");
    expect(normalizeFormulaLocaleId("zh")).toBe("zh-CN");
    expect(normalizeFormulaLocaleId("zh-Hans")).toBe("zh-CN");
    expect(normalizeFormulaLocaleId("zh-Hant")).toBe("zh-TW");
    expect(normalizeFormulaLocaleId("zh-Hant-u-nu-latn")).toBe("zh-TW");
    expect(normalizeFormulaLocaleId("zh-HK")).toBe("zh-TW");
    expect(normalizeFormulaLocaleId("zh-TW")).toBe("zh-TW");

    // POSIX "C locale" aliases.
    expect(normalizeFormulaLocaleId("C")).toBe("en-US");
    expect(normalizeFormulaLocaleId("C.UTF-8")).toBe("en-US");
    expect(normalizeFormulaLocaleId("POSIX")).toBe("en-US");

    // Unknown locales stay unknown.
    expect(normalizeFormulaLocaleId("pt-BR")).toBeNull();
    expect(normalizeFormulaLocaleId("it-IT")).toBeNull();
    expect(normalizeFormulaLocaleId("")).toBeNull();
  });
});
