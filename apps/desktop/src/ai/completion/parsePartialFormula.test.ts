import { afterEach, describe, expect, it, vi } from "vitest";

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

  it("canonicalizes localized function names case-insensitively (zählenwenn -> COUNTIF)", async () => {
    setLocale("de-DE");

    const engine = {
      parseFormulaPartial: async () => {
        // Intentionally lowercase to ensure we match Rust's Unicode-aware case folding.
        return { context: { function: { name: "zählenwenn", argIndex: 0 } } };
      },
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 1000,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=zählenwenn(A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("COUNTIF");
    expect(result.expectingRange).toBe(true);
  });

  it("canonicalizes localized function names even when falling back to the JS parser", async () => {
    setLocale("de-DE");

    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMME(A1;";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.argIndex).toBe(1);
    expect(result.expectingRange).toBe(true);
  });

  it("does not treat decimal commas as argument separators in semicolon locales", async () => {
    setLocale("de-DE");

    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMME(1,";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    // The de-DE locale uses ';' for arguments and ',' for decimals.
    expect(result.argIndex).toBe(0);
    expect(result.currentArg?.text).toBe("1,");
    expect(result.functionName).toBe("SUM");
  });

  it("does not treat separators inside structured refs with escaped brackets as function args (semicolon locales)", async () => {
    setLocale("de-DE");

    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    // Column name is literally `A]B;USD`, encoded as `A]]B;USD` inside the structured ref item.
    // The semicolon inside the structured ref must not be treated as a function argument separator.
    const input = "=SUMME(Table1[[#Headers],[A]]B;USD]]; A";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.argIndex).toBe(1);
    expect(result.currentArg?.text).toBe("A");
  });

  it("does not treat separators inside external workbook prefixes as function args (semicolon locales)", async () => {
    setLocale("de-DE");

    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=SUMME([A1[Name.xlsx]Sheet1!A1; 1";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    expect(result.functionName).toBe("SUM");
    expect(result.argIndex).toBe(1);
    expect(result.currentArg?.text).toBe("1");
  });

  it("does not treat decimal commas as argument separators for A1-like function names (LOG10)", async () => {
    setLocale("de-DE");

    const parser = createLocaleAwarePartialFormulaParser({});
    const fnRegistry = new FunctionRegistry();

    const input = "=LOG10(1,";
    const result = await parser(input, input.length, fnRegistry);

    expect(result.isFormula).toBe(true);
    expect(result.inFunctionCall).toBe(true);
    // LOG10 resembles a cell reference (LOG + 10), but should still be parsed as a function call.
    expect(result.functionName).toBe("LOG10");
    // The de-DE locale uses ';' for arguments and ',' for decimals.
    expect(result.argIndex).toBe(0);
    expect(result.currentArg?.text).toBe("1,");
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

  it("falls back to the JS parser when the engine parser times out", async () => {
    setLocale("en-US");

    const engine = {
      // Simulate a hung worker / slow engine: never resolves.
      parseFormulaPartial: async () => new Promise(() => {}),
    };

    const parser = createLocaleAwarePartialFormulaParser({
      getEngineClient: () => engine,
      timeoutMs: 5,
    });
    const fnRegistry = new FunctionRegistry();

    const input = "=SUM(A";
    const expected = parsePartialFormulaFallback(input, input.length, fnRegistry);

    vi.useFakeTimers();
    try {
      const pending = parser(input, input.length, fnRegistry);
      await vi.advanceTimersByTimeAsync(5);
      await expect(pending).resolves.toEqual(expected);
    } finally {
      vi.useRealTimers();
    }
  });

  it("prefers document.lang over the i18n locale when choosing the formula locale", async () => {
    // The desktop shell may set `<html lang="...">` without calling `setLocale()`.
    // Ensure the parser follows the document locale so localized formula UX stays consistent.
    setLocale("en-US");
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = { documentElement: { lang: "de-DE" } };

    try {
      const parser = createLocaleAwarePartialFormulaParser({});
      const fnRegistry = new FunctionRegistry();
      const input = "=SUMME(1,";
      const result = await parser(input, input.length, fnRegistry);

      expect(result.isFormula).toBe(true);
      expect(result.inFunctionCall).toBe(true);
      // `,` is a decimal separator in de-DE, so argIndex should remain 0.
      expect(result.argIndex).toBe(0);
      expect(result.currentArg?.text).toBe("1,");
      expect(result.functionName).toBe("SUM");
    } finally {
      (globalThis as any).document = prevDocument;
    }
  });
});
