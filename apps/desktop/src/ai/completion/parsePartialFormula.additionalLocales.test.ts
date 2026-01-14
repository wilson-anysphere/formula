import { describe, expect, it, vi } from "vitest";

import { FunctionRegistry } from "@formula/ai-completion";

// The desktop UI currently only ships `en-US`, `de-DE`, and `ar` translations, but the
// WASM formula engine supports additional locales (e.g. `fr-FR`, `es-ES`) for parsing.
//
// `createLocaleAwarePartialFormulaParser` relies on `getLocale()` to select the locale,
// so mock it here to ensure the localized->canonical function mapping tables remain correct.
let currentLocale = "fr-FR";
vi.mock("../../i18n/index.js", () => ({
  getLocale: () => currentLocale,
}));

describe("createLocaleAwarePartialFormulaParser (engine-supported locales)", () => {
  it("canonicalizes fr-FR function names (SOMME -> SUM)", async () => {
    currentLocale = "fr-FR";
    const { createLocaleAwarePartialFormulaParser } = await import("./parsePartialFormula.js");
    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=SOMME(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.expectingRange).toBe(true);
  });

  it("canonicalizes fr-FR dotted function names (NB.SI -> COUNTIF)", async () => {
    currentLocale = "fr-FR";
    const { createLocaleAwarePartialFormulaParser } = await import("./parsePartialFormula.js");
    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=NB.SI(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("COUNTIF");
    expect(result.expectingRange).toBe(true);
  });

  it("canonicalizes es-ES function names (SUMA -> SUM)", async () => {
    currentLocale = "es-ES";
    const { createLocaleAwarePartialFormulaParser } = await import("./parsePartialFormula.js");
    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMA(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.expectingRange).toBe(true);
  });

  it("treats ',' as a decimal separator (not an arg separator) in semicolon locales", async () => {
    const { createLocaleAwarePartialFormulaParser } = await import("./parsePartialFormula.js");
    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    currentLocale = "fr-FR";
    const frInput = "=SOMME(1,";
    const frResult = await parser(frInput, frInput.length, fnRegistry);
    expect(frResult.argIndex).toBe(0);
    expect(frResult.currentArg?.text).toBe("1,");
    expect(frResult.functionName).toBe("SUM");

    currentLocale = "es-ES";
    const esInput = "=SUMA(1,";
    const esResult = await parser(esInput, esInput.length, fnRegistry);
    expect(esResult.argIndex).toBe(0);
    expect(esResult.currentArg?.text).toBe("1,");
    expect(esResult.functionName).toBe("SUM");
  });

  it("normalizes POSIX/variant locale IDs to the supported engine locale (de_DE.UTF-8 → de-DE)", async () => {
    const { createLocaleAwarePartialFormulaParser } = await import("./parsePartialFormula.js");
    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    currentLocale = "de_DE.UTF-8";
    const input = "=SUMME(1,";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    // de-DE uses semicolons, so the trailing comma is treated as a decimal separator.
    expect(result.argIndex).toBe(0);
    expect(result.currentArg?.text).toBe("1,");
    // Localized function names are canonicalized for metadata lookup.
    expect(result.functionName).toBe("SUM");
  });
});
