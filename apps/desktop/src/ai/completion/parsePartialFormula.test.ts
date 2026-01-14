import { afterEach, describe, expect, it, vi } from "vitest";

import { FunctionRegistry, parsePartialFormula as parsePartialFormulaFallback } from "@formula/ai-completion";

import { setLocale } from "../../i18n/index.js";
import {
  createLocaleAwareFunctionRegistry,
  createLocaleAwarePartialFormulaParser,
  createLocaleAwareStarterFunctions,
} from "./parsePartialFormula.js";

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
    // Simulate an engine version mismatch: the host thinks a locale is supported, but the
    // engine rejects it. The wrapper should cache this signal and avoid retrying on every keypress.
    setLocale("de-DE");

    let calls = 0;
    const engine = {
      parseFormulaPartial: async () => {
        calls += 1;
        throw new Error("unknown localeId: de-DE. Supported locale ids: en-US, de-DE");
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

  it("falls back to en-US semantics when document.lang is an unsupported locale (pt-BR)", async () => {
    // Some hosts report locales that the formula engine doesn't support yet (e.g. pt-BR).
    // Formula UX should match the engine's behavior (treat unsupported locales as en-US)
    // so arg separators and parsing remain consistent.
    setLocale("en-US");
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = { documentElement: { lang: "pt-BR" } };

    try {
      const parser = createLocaleAwarePartialFormulaParser({});
      const fnRegistry = new FunctionRegistry();

      const commaInput = "=SUM(1,";
      const commaResult = await parser(commaInput, commaInput.length, fnRegistry);
      expect(commaResult.isFormula).toBe(true);
      expect(commaResult.inFunctionCall).toBe(true);
      // In the en-US fallback, ',' is treated as an argument separator.
      expect(commaResult.argIndex).toBe(1);
      expect(commaResult.currentArg?.text).toBe("");
      expect(commaResult.functionName).toBe("SUM");

      const semiInput = "=SUM(1;";
      const semiResult = await parser(semiInput, semiInput.length, fnRegistry);
      // In the en-US fallback, ';' is treated as plain text (not an argument separator).
      expect(semiResult.argIndex).toBe(0);
      expect(semiResult.currentArg?.text).toBe("1;");
    } finally {
      (globalThis as any).document = prevDocument;
    }
  });

  it("does not treat semicolons inside array literals as argument separators in comma locales (en-US)", async () => {
    setLocale("en-US");
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = { documentElement: { lang: "en-US" } };

    try {
      const parser = createLocaleAwarePartialFormulaParser({});
      const fnRegistry = new FunctionRegistry();

      // In en-US, semicolons are used as array row separators (inside `{...}`), not function arg separators.
      // Ensure they don't affect argument indexing.
      const input = "=SUM({1;2},";
      const result = await parser(input, input.length, fnRegistry);

      expect(result.isFormula).toBe(true);
      expect(result.inFunctionCall).toBe(true);
      expect(result.functionName).toBe("SUM");
      expect(result.argIndex).toBe(1);
      expect(result.currentArg?.text).toBe("");
    } finally {
      (globalThis as any).document = prevDocument;
    }
  });

  it("respects getLocaleId() overrides over document/i18n locale when choosing the formula locale", async () => {
    setLocale("en-US");
    const prevDocument = (globalThis as any).document;
    (globalThis as any).document = { documentElement: { lang: "en-US" } };

    try {
      const parser = createLocaleAwarePartialFormulaParser({ getLocaleId: () => "de-DE" });
      const fnRegistry = new FunctionRegistry();
      const input = "=SUMME(1,";
      const result = await parser(input, input.length, fnRegistry);

      expect(result.isFormula).toBe(true);
      expect(result.inFunctionCall).toBe(true);
      // In de-DE, ',' is a decimal separator, not an argument separator.
      expect(result.argIndex).toBe(0);
      expect(result.currentArg?.text).toBe("1,");
      // Localized function name should be canonicalized.
      expect(result.functionName).toBe("SUM");
    } finally {
      (globalThis as any).document = prevDocument;
    }
  });

  it("createLocaleAwareStarterFunctions respects getLocaleId() overrides", () => {
    setLocale("en-US");
    const starters = createLocaleAwareStarterFunctions({ getLocaleId: () => "de-DE" });
    expect(starters()[0]).toBe("SUMME(");
  });

  it("createLocaleAwareFunctionRegistry respects getLocaleId() overrides", () => {
    setLocale("en-US");

    const deRegistry = createLocaleAwareFunctionRegistry({ getLocaleId: () => "de-DE" });
    const deMatches = deRegistry.search("SU", { limit: 50 });
    expect(deMatches.some((m) => m.name === "SUMME")).toBe(true);

    const enRegistry = createLocaleAwareFunctionRegistry({ getLocaleId: () => "en-US" });
    const enMatches = enRegistry.search("SU", { limit: 50 });
    expect(enMatches.some((m) => m.name === "SUMME")).toBe(false);
  });
});
