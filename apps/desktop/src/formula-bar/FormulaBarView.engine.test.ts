/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import type { EngineClient, FormulaPartialLexResult, FormulaPartialParseResult, FormulaToken } from "@formula/engine";

import { FormulaBarView } from "./FormulaBarView.js";

async function flushTooling(): Promise<void> {
  const nextFrame = () =>
    new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(resolve, 0);
      }
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
});
