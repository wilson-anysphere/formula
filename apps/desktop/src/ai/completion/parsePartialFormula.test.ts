import { afterEach, describe, expect, it } from "vitest";

import { FunctionRegistry, parsePartialFormula as parsePartialFormulaFallback } from "@formula/ai-completion";

import { setLocale } from "../../i18n/index.js";
import { createLocaleAwarePartialFormulaParser } from "./parsePartialFormula.js";

describe("createLocaleAwarePartialFormulaParser", () => {
  afterEach(() => {
    setLocale("en-US");
  });

  it("uses the engine parser for semicolon locales (argIndex + expectingRange)", async () => {
    setLocale("de-DE");

    /** @type {any[]} */
    const calls = [];
    const engine = {
      parseFormulaPartial: async (formula: string, cursor: number, options: any) => {
        calls.push({ formula, cursor, options });
        const prefix = formula.slice(0, cursor);
        // Minimal "engine-like" result: count top-level semicolons inside the SUM(...) call.
        const argIndex = prefix.includes(";") ? 1 : 0;
        return { context: { function: { name: "SUMME", argIndex } } };
      },
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 1000,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMME(A1;";
    const result = await parser(input, input.length, fnRegistry);

    expect(calls.length).toBe(1);
    expect(calls[0]?.options?.localeId).toBe("de-DE");

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    // Adapter canonicalizes localized function names for completion metadata lookup.
    expect(result.functionName).toBe("SUM");
    expect(result.argIndex).toBe(1);
    expect(result.expectingRange).toBe(true);
    expect(result.currentArg?.text).toBe("");
    expect(result.currentArg?.start).toBe(input.length);
  });

  it("canonicalizes localized function names for signature/range metadata (SUMME -> SUM)", async () => {
    setLocale("de-DE");

    const engine = {
      parseFormulaPartial: async () => {
        return { context: { function: { name: "SUMME", argIndex: 0 } } };
      },
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 1000,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMME(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    // Completion engine metadata is keyed by canonical (English) names.
    expect(result.functionName).toBe("SUM");
    expect(result.expectingRange).toBe(true);
  });

  it("falls back to the JS parser when the engine throws", async () => {
    setLocale("en-US");

    let calls = 0;
    const engine = {
      parseFormulaPartial: async () => {
        calls += 1;
        throw new Error("engine unavailable");
      },
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 1000,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUM(A1,";
    const expected = parsePartialFormulaFallback(input, input.length, fnRegistry);
    const result = await parser(input, input.length, fnRegistry);

    expect(calls).toBe(1);
    expect(result).toEqual(expected);
  });

  it("caches unsupported localeIds to avoid repeated engine RPC failures", async () => {
    setLocale("ar");

    let calls = 0;
    const engine = {
      parseFormulaPartial: async () => {
        calls += 1;
        throw new Error("unknown localeId: ar");
      },
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 1000,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUM(A1,";
    await parser(input, input.length, fnRegistry);
    await parser(input, input.length, fnRegistry);

    expect(calls).toBe(1);
  });
});
