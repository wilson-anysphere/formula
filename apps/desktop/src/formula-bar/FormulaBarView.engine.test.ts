/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import type { EngineClient, FormulaPartialLexResult, FormulaPartialParseResult, FormulaToken } from "@formula/engine";

import { FormulaBarView } from "./FormulaBarView.js";

async function flushTooling(): Promise<void> {
  const nextFrame = (): Promise<void> =>
    new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") requestAnimationFrame(() => resolve());
      else setTimeout(resolve, 0);
    });

  // Frame 1: flush the scheduled render + scheduleEngineTooling tick.
  await nextFrame();
  // Frame 2: flush the render scheduled after the async engine result is applied.
  await nextFrame();
  // Frame 3: best-effort extra tick for environments where RAF is emulated via timers.
  await nextFrame();
}

describe("FormulaBarView WASM editor tooling integration", () => {
  it("wraps the engine parse error span in a dedicated error class", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const draft = "=1+";

    const tokens: FormulaToken[] = [
      { kind: "Number", span: { start: 1, end: 2 }, value: "1" },
      { kind: "Plus", span: { start: 2, end: 3 } },
      { kind: "Eof", span: { start: 3, end: 3 } },
    ];

    const lexResult: FormulaPartialLexResult = { tokens, error: null };
    const parseResult: FormulaPartialParseResult = {
      ast: null,
      error: { message: "Unexpected token", span: { start: 2, end: 3 } },
      context: { function: null },
    };

    const engine = {
      lexFormulaPartial: vi.fn(async () => lexResult),
      parseFormulaPartial: vi.fn(async () => parseResult),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {} },
      { getWasmEngine: () => engine, getLocaleId: () => "en-US", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    view.textarea.value = draft;
    view.textarea.setSelectionRange(draft.length, draft.length);
    view.textarea.dispatchEvent(new Event("input"));

    await flushTooling();

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const eqSpan = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="operator"]') ?? []).find(
      (el) => el.textContent === "=",
    );
    expect(eqSpan).toBeTruthy();
    const errorEl = highlight?.querySelector<HTMLElement>(".formula-bar-token--error");
    expect(errorEl).toBeTruthy();
    expect(errorEl?.textContent).toBe("+");
    expect(host.querySelector<HTMLElement>('[data-testid="formula-hint"]')?.textContent).toContain("Unexpected token");

    host.remove();
  });

  it("reuses the cached lex result when only the cursor moves", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const draft = "=SUM(1, 2)";

    const tokens: FormulaToken[] = [
      { kind: "Eq", span: { start: 0, end: 1 } },
      { kind: "Ident", span: { start: 1, end: 4 }, value: "SUM" },
      { kind: "LParen", span: { start: 4, end: 5 } },
      { kind: "Number", span: { start: 5, end: 6 }, value: "1" },
      { kind: "ArgSep", span: { start: 6, end: 7 } },
      { kind: "Whitespace", span: { start: 7, end: 8 }, value: " " },
      { kind: "Number", span: { start: 8, end: 9 }, value: "2" },
      { kind: "RParen", span: { start: 9, end: 10 } },
      { kind: "Eof", span: { start: 10, end: 10 } },
    ];

    const lexResult: FormulaPartialLexResult = { tokens, error: null };

    const engine = {
      lexFormulaPartial: vi.fn(async () => lexResult),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: { name: "SUM", argIndex: 0 } } })),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {} },
      { getWasmEngine: () => engine, getLocaleId: () => "en-US", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    view.textarea.value = draft;
    view.textarea.setSelectionRange(draft.length, draft.length);
    view.textarea.dispatchEvent(new Event("input"));

    await flushTooling();

    expect(engine.lexFormulaPartial).toHaveBeenCalledTimes(1);
    expect(engine.parseFormulaPartial).toHaveBeenCalledTimes(1);

    // Cursor move within the same draft should trigger a new parse call but reuse the cached lex result.
    const cursorArg0 = draft.indexOf("1") + 1;
    view.textarea.setSelectionRange(cursorArg0, cursorArg0);
    view.textarea.dispatchEvent(new Event("select"));

    await flushTooling();

    expect(engine.lexFormulaPartial).toHaveBeenCalledTimes(1);
    expect(engine.parseFormulaPartial).toHaveBeenCalledTimes(2);

    host.remove();
  });

  it("updates the function hint using engine parse context", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    // Use a locale-sensitive argument separator to ensure we're actually using the
    // engine parse context (the local fallback only understands commas).
    const draft = "=SUM(1; 2)";
    const sepIndex = draft.indexOf(";");

    const tokens: FormulaToken[] = [
      { kind: "Ident", span: { start: 1, end: 4 }, value: "SUM" },
      { kind: "LParen", span: { start: 4, end: 5 } },
      { kind: "Number", span: { start: 5, end: 6 }, value: "1" },
      { kind: "ArgSep", span: { start: 6, end: 7 } },
      { kind: "Whitespace", span: { start: 7, end: 8 }, value: " " },
      { kind: "Number", span: { start: 8, end: 9 }, value: "2" },
      { kind: "RParen", span: { start: 9, end: 10 } },
      { kind: "Eof", span: { start: 10, end: 10 } },
    ];

    const lexResult: FormulaPartialLexResult = { tokens, error: null };

    const engine = {
      lexFormulaPartial: vi.fn(async () => lexResult),
      parseFormulaPartial: vi.fn(async (formula: string, cursor?: number) => {
        const argIndex = cursor != null && cursor <= sepIndex ? 0 : 1;
        return {
          ast: null,
          error: null,
          context: { function: { name: "SUM", argIndex } },
        } satisfies FormulaPartialParseResult;
      }),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {} },
      { getWasmEngine: () => engine, getLocaleId: () => "de-DE", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const hintEl = () => host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
    const activeParamText = () =>
      hintEl()?.querySelector<HTMLElement>(".formula-bar-hint-token--paramActive")?.textContent ?? "";
    const signatureText = () => hintEl()?.querySelector<HTMLElement>(".formula-bar-hint-signature")?.textContent ?? "";

    // Cursor within the first argument.
    const cursorArg0 = draft.indexOf("1") + 1;
    view.textarea.value = draft;
    view.textarea.setSelectionRange(cursorArg0, cursorArg0);
    view.textarea.dispatchEvent(new Event("input"));
    await flushTooling();
    expect(activeParamText()).toBe("number1");
    expect(signatureText()).toContain("; ");

    // Cursor within the second argument.
    const cursorArg1 = draft.indexOf("2") + 1;
    view.textarea.setSelectionRange(cursorArg1, cursorArg1);
    view.textarea.dispatchEvent(new Event("select"));
    await flushTooling();
    expect(activeParamText()).toBe("[number2]");
    expect(signatureText()).toContain("; ");

    host.remove();
  });

  it("normalizes unsupported locale IDs to en-US so engine tooling stays available", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const draft = "=SUM(1, 2)";

    const tokens: FormulaToken[] = [
      { kind: "Eq", span: { start: 0, end: 1 } },
      { kind: "Ident", span: { start: 1, end: 4 }, value: "SUM" },
      { kind: "LParen", span: { start: 4, end: 5 } },
      { kind: "Number", span: { start: 5, end: 6 }, value: "1" },
      { kind: "ArgSep", span: { start: 6, end: 7 } },
      { kind: "Whitespace", span: { start: 7, end: 8 }, value: " " },
      { kind: "Number", span: { start: 8, end: 9 }, value: "2" },
      { kind: "RParen", span: { start: 9, end: 10 } },
      { kind: "Eof", span: { start: 10, end: 10 } },
    ];

    const lexResult: FormulaPartialLexResult = { tokens, error: null };
    const parseResult: FormulaPartialParseResult = { ast: null, error: null, context: { function: null } };

    const engine = {
      lexFormulaPartial: vi.fn(async () => lexResult),
      parseFormulaPartial: vi.fn(async () => parseResult),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {} },
      // pt-BR is not currently supported by the formula engine's locale registry. The view should
      // fall back to en-US so tooling requests don't hard-fail.
      { getWasmEngine: () => engine, getLocaleId: () => "pt-BR", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    view.textarea.value = draft;
    view.textarea.setSelectionRange(draft.length, draft.length);
    view.textarea.dispatchEvent(new Event("input"));

    await flushTooling();

    expect(engine.lexFormulaPartial).toHaveBeenCalledTimes(1);
    const lexOptions = (engine.lexFormulaPartial as any).mock.calls[0]?.[1];
    expect(lexOptions?.localeId).toBe("en-US");

    expect(engine.parseFormulaPartial).toHaveBeenCalledTimes(1);
    const parseOptions = (engine.parseFormulaPartial as any).mock.calls[0]?.[2];
    expect(parseOptions?.localeId).toBe("en-US");

    host.remove();
  });

  it("falls back to the local tokenizer/highlighter when the engine is absent", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    const draft = "=ROUND(1, 2)";
    const cursor = draft.indexOf("1") + 1;
    view.textarea.value = draft;
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("input"));
    await flushTooling();

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const fn = highlight?.querySelector<HTMLElement>('span[data-kind="function"]');
    expect(fn?.textContent).toBe("ROUND");

    const hint = host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
    expect(hint?.textContent).toContain("ROUND(");

    host.remove();
  });

  it("uses locale-aware argument separators in fallback function hints", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });
      view.focus({ cursor: "end" });

      const draft = "=SUM(1; 2)";
      const cursorArg1 = draft.indexOf("2") + 1;
      view.textarea.value = draft;
      view.textarea.setSelectionRange(cursorArg1, cursorArg1);
      view.textarea.dispatchEvent(new Event("input"));

      await flushTooling();

      const hintEl = host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
      const signatureText = hintEl?.querySelector<HTMLElement>(".formula-bar-hint-signature")?.textContent ?? "";
      expect(signatureText).toContain("; ");
    } finally {
      document.documentElement.lang = prevLang;
      host.remove();
    }
  });

  it("ignores stale out-of-order engine tooling responses", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    type Deferred<T> = { promise: Promise<T>; resolve: (value: T) => void; reject: (err: unknown) => void };
    const defer = <T,>(): Deferred<T> => {
      let resolve!: (value: T) => void;
      let reject!: (err: unknown) => void;
      const promise = new Promise<T>((res, rej) => {
        resolve = res;
        reject = rej;
      });
      return { promise, resolve, reject };
    };

    const lexCalls: Array<{ formula: string; deferred: Deferred<FormulaPartialLexResult> }> = [];
    const parseCalls: Array<{ formula: string; deferred: Deferred<FormulaPartialParseResult> }> = [];

    const engine = {
      lexFormulaPartial: vi.fn((formula: string) => {
        const deferred = defer<FormulaPartialLexResult>();
        lexCalls.push({ formula, deferred });
        return deferred.promise;
      }),
      parseFormulaPartial: vi.fn((formula: string) => {
        const deferred = defer<FormulaPartialParseResult>();
        parseCalls.push({ formula, deferred });
        return deferred.promise;
      }),
    } as unknown as EngineClient;

    const view = new FormulaBarView(host, { onCommit: () => {} }, { getWasmEngine: () => engine, getLocaleId: () => "en-US" });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const nextFrame = (): Promise<void> =>
      new Promise<void>((resolve) => {
        if (typeof requestAnimationFrame === "function") requestAnimationFrame(() => resolve());
        else setTimeout(resolve, 0);
      });

    // Draft 1 -> schedule tooling and let the engine calls start (but keep them pending).
    const draft1 = "=1+";
    view.textarea.value = draft1;
    view.textarea.setSelectionRange(draft1.length, draft1.length);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();
    expect(lexCalls.length).toBe(1);
    expect(parseCalls.length).toBe(1);
    expect(lexCalls[0]?.formula).toBe(draft1);

    // Draft 2 -> start a second tooling request (also pending).
    const draft2 = "=1+2";
    view.textarea.value = draft2;
    view.textarea.setSelectionRange(draft2.length, draft2.length);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();
    expect(lexCalls.length).toBe(2);
    expect(parseCalls.length).toBe(2);
    expect(lexCalls[1]?.formula).toBe(draft2);

    // Resolve the *second* call first (valid formula, no syntax error).
    lexCalls[1]!.deferred.resolve({
      tokens: [
        { kind: "Number", span: { start: 1, end: 2 }, value: "1" },
        { kind: "Plus", span: { start: 2, end: 3 } },
        { kind: "Number", span: { start: 3, end: 4 }, value: "2" },
        { kind: "Eof", span: { start: 4, end: 4 } },
      ],
      error: null,
    });
    parseCalls[1]!.deferred.resolve({ ast: null, error: null, context: { function: null } });
    await flushTooling();

    const highlight = () => host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight()?.textContent).toBe(draft2);
    expect(highlight()?.querySelector(".formula-bar-token--error")).toBeNull();

    // Now resolve the *first* call late (it has a syntax error). This must NOT override draft2.
    lexCalls[0]!.deferred.resolve({
      tokens: [
        { kind: "Number", span: { start: 1, end: 2 }, value: "1" },
        { kind: "Plus", span: { start: 2, end: 3 } },
        { kind: "Eof", span: { start: 3, end: 3 } },
      ],
      error: null,
    });
    parseCalls[0]!.deferred.resolve({
      ast: null,
      error: { message: "Unexpected token", span: { start: 2, end: 3 } },
      context: { function: null },
    });
    await flushTooling();

    expect(highlight()?.textContent).toBe(draft2);
    expect(highlight()?.querySelector(".formula-bar-token--error")).toBeNull();

    host.remove();
  });

  it("does not invoke engine tooling after canceling before the scheduled flush runs", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const engine = {
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {}, onCancel: () => {} },
      { getWasmEngine: () => engine, getLocaleId: () => "en-US", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    view.textarea.value = "=1+2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    // Cancel immediately, before the RAF/timer tooling flush runs.
    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));

    await flushTooling();

    expect(engine.lexFormulaPartial).toHaveBeenCalledTimes(0);
    expect(engine.parseFormulaPartial).toHaveBeenCalledTimes(0);

    host.remove();
  });

  it("aborts in-flight engine tooling when canceling", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const signals: AbortSignal[] = [];
    const abortable = (rpc: any) =>
      new Promise((_, reject) => {
        const signal = rpc?.signal as AbortSignal | undefined;
        if (signal) {
          signals.push(signal);
          if (signal.aborted) {
            reject(new Error("aborted"));
            return;
          }
          signal.addEventListener(
            "abort",
            () => {
              reject(new Error("aborted"));
            },
            { once: true },
          );
        }
      });

    const engine = {
      lexFormulaPartial: vi.fn((_formula: string, _opts: any, rpc: any) => abortable(rpc) as any),
      parseFormulaPartial: vi.fn((_formula: string, _cursor: number, _opts: any, rpc: any) => abortable(rpc) as any),
    } as unknown as EngineClient;

    const view = new FormulaBarView(
      host,
      { onCommit: () => {}, onCancel: () => {} },
      { getWasmEngine: () => engine, getLocaleId: () => "en-US", referenceStyle: "A1" },
    );

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });
    view.textarea.value = "=1+2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    // Allow the scheduled tooling flush to run and start the async engine requests.
    await new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") requestAnimationFrame(() => resolve());
      else setTimeout(() => resolve(), 0);
    });

    expect(engine.lexFormulaPartial).toHaveBeenCalledTimes(1);
    expect(engine.parseFormulaPartial).toHaveBeenCalledTimes(1);
    expect(signals.length).toBeGreaterThan(0);
    expect(signals.every((s) => !s.aborted)).toBe(true);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));

    expect(signals.every((s) => s.aborted)).toBe(true);

    // Let the abort rejection settle in the background (FormulaBarView swallows it).
    await Promise.resolve();

    host.remove();
  });

  it("does not invoke engine tooling for non-formula text", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const engine = {
      lexFormulaPartial: vi.fn(async () => ({ tokens: [], error: null })),
      parseFormulaPartial: vi.fn(async () => ({ ast: null, error: null, context: { function: null } })),
    } as unknown as EngineClient;

    const view = new FormulaBarView(host, { onCommit: () => {} }, { getWasmEngine: () => engine, getLocaleId: () => "en-US" });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.textarea.value = "hello";
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("input"));

    await flushTooling();

    expect(engine.lexFormulaPartial).not.toHaveBeenCalled();
    expect(engine.parseFormulaPartial).not.toHaveBeenCalled();

    host.remove();
  });
});
