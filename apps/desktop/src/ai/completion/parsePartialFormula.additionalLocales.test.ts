import { describe, expect, it } from "vitest";

import { FunctionRegistry } from "@formula/ai-completion";

import { createLocaleAwarePartialFormulaParser } from "./parsePartialFormula.js";

describe("createLocaleAwarePartialFormulaParser (engine-supported locales)", () => {
  it("canonicalizes fr-FR function names (SOMME -> SUM)", async () => {
    const parser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "fr-FR" });
    const fnRegistry = new FunctionRegistry();

    const input = "=SOMME(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.expectingRange).toBe(true);
  });

  it("canonicalizes fr-FR dotted function names (NB.SI -> COUNTIF)", async () => {
    const parser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "fr-FR" });
    const fnRegistry = new FunctionRegistry();

    const input = "=NB.SI(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("COUNTIF");
    expect(result.expectingRange).toBe(true);
  });

  it("canonicalizes es-ES function names (SUMA -> SUM)", async () => {
    const parser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "es-ES" });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMA(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.expectingRange).toBe(true);
  });

  it("treats ',' as a decimal separator (not an arg separator) in semicolon locales", async () => {
    const fnRegistry = new FunctionRegistry();

    const frParser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "fr-FR" });
    const frInput = "=SOMME(1,";
    const frResult = await frParser(frInput, frInput.length, fnRegistry);
    expect(frResult.argIndex).toBe(0);
    expect(frResult.currentArg?.text).toBe("1,");
    expect(frResult.functionName).toBe("SUM");

    const esParser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "es-ES" });
    const esInput = "=SUMA(1,";
    const esResult = await esParser(esInput, esInput.length, fnRegistry);
    expect(esResult.argIndex).toBe(0);
    expect(esResult.currentArg?.text).toBe("1,");
    expect(esResult.functionName).toBe("SUM");
  });

  it("normalizes POSIX/variant locale IDs to the supported engine locale (de_DE.UTF-8 → de-DE)", async () => {
    const parser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "de_DE.UTF-8" });
    const fnRegistry = new FunctionRegistry();

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
