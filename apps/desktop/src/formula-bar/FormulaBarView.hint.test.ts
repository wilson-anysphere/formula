/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

async function nextFrame(): Promise<void> {
  await new Promise<void>((resolve) => {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => resolve());
    } else {
      setTimeout(() => resolve(), 0);
    }
  });
}

function getHintEl(host: HTMLElement): HTMLElement {
  const hint = host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
  if (!hint) throw new Error("Expected formula hint element");
  return hint;
}

function getActiveParamText(host: HTMLElement): string | null {
  const hint = getHintEl(host);
  return hint.querySelector<HTMLElement>(".formula-bar-hint-token--paramActive")?.textContent ?? null;
}

function getSignatureName(host: HTMLElement): string | null {
  const hint = getHintEl(host);
  return hint.querySelector<HTMLElement>(".formula-bar-hint-token--name")?.textContent ?? null;
}

describe("FormulaBarView function hint UI", () => {
  it("does not show a function hint when not editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "=ROUND(1, 2)", value: null });

    const hint = getHintEl(host);
    expect(hint.querySelector(".formula-bar-hint-panel")).toBeNull();
    expect(hint.textContent).toBe("");

    host.remove();
  });

  it("treats whitespace between the function name and '(' as a function call (Excel-style)", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=IF (A1>0,1,2)";

    const inFirstArg = view.textarea.value.indexOf(">") + 1;
    view.textarea.setSelectionRange(inFirstArg, inFirstArg);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();
    expect(getSignatureName(host)).toBe("IF(");
    expect(getActiveParamText(host)).toBe("logical_test");

    host.remove();
  });

  it("does not treat decimal commas as argument separators in semicolon locales", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=ROUND(1,2; 0)";

      const inFirstArg = view.textarea.value.indexOf("1,2") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("ROUND(");
      // Cursor in the first argument; the comma is a decimal separator (1,2), not an argument separator.
      expect(getActiveParamText(host)).toBe("number");

      const inSecondArg = view.textarea.value.indexOf("0") + 1;
      view.textarea.setSelectionRange(inSecondArg, inSecondArg);
      view.textarea.dispatchEvent(new Event("select"));
      await nextFrame();
      expect(getActiveParamText(host)).toBe("num_digits");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("normalizes POSIX locale IDs when inferring argument separators (de_DE.UTF-8)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de_DE.UTF-8";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=ROUND(1,2; 0)";

      const inFirstArg = view.textarea.value.indexOf("1,2") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("ROUND(");
      // Cursor in the first argument; the comma is a decimal separator (1,2), not an argument separator.
      expect(getActiveParamText(host)).toBe("number");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("treats unsupported locale IDs as en-US when inferring argument separators (pt-BR)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "pt-BR";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      // In the current engine, pt-BR is unsupported and formula punctuation falls back to en-US.
      // That means commas are treated as argument separators (not decimal separators).
      view.textarea.value = "=ROUND(1,2)";

      const inFirstArg = view.textarea.value.indexOf("1") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();
      expect(getSignatureName(host)).toBe("ROUND(");
      expect(getActiveParamText(host)).toBe("number");

      const inSecondArg = view.textarea.value.lastIndexOf("2") + 1;
      view.textarea.setSelectionRange(inSecondArg, inSecondArg);
      view.textarea.dispatchEvent(new Event("select"));
      await nextFrame();
      expect(getActiveParamText(host)).toBe("num_digits");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("normalizes BCP-47 variants when inferring argument separators (de-CH-1996)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-CH-1996";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=ROUND(1,2; 0)";

      const inFirstArg = view.textarea.value.indexOf("1,2") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("ROUND(");
      // Even though `de-CH` uses different numeric formatting, the formula engine currently treats
      // it as `de-DE` for formula parsing (decimal comma + `;` list separator). Ensure the hint
      // logic stays consistent with that behavior.
      expect(getActiveParamText(host)).toBe("number");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("normalizes language-only locale IDs when resolving localized function names (de SUMME)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=SUMME(A1; B1)";

      const inFirstArg = view.textarea.value.indexOf("A1") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("SUMME(");
      expect(getActiveParamText(host)).toBe("number1");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("shows function hints for localized function names (de-DE SUMME)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=SUMME(A1; B1)";

      const inFirstArg = view.textarea.value.indexOf("A1") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("SUMME(");
      expect(getActiveParamText(host)).toBe("number1");

      const inSecondArg = view.textarea.value.indexOf("B1") + 1;
      view.textarea.setSelectionRange(inSecondArg, inSecondArg);
      view.textarea.dispatchEvent(new Event("select"));
      await nextFrame();

      expect(getActiveParamText(host)).toBe("[number2]");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("shows function hints for localized function names with non-ASCII letters (de-DE ZÄHLENWENN)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = '=ZÄHLENWENN(A1:A3; ">0")';

      const inFirstArg = view.textarea.value.indexOf("A1") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("ZÄHLENWENN(");
      expect(getActiveParamText(host)).toBe("range");

      const inSecondArg = view.textarea.value.indexOf('">0"') + 2;
      view.textarea.setSelectionRange(inSecondArg, inSecondArg);
      view.textarea.dispatchEvent(new Event("select"));
      await nextFrame();

      expect(getActiveParamText(host)).toBe("criteria");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("shows function hints for dotted localized function names (es-ES CONTAR.SI)", async () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "es-ES";

    try {
      const host = document.createElement("div");
      document.body.appendChild(host);

      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = '=CONTAR.SI(A1:A3; ">0")';

      const inFirstArg = view.textarea.value.indexOf("A1") + 1;
      view.textarea.setSelectionRange(inFirstArg, inFirstArg);
      view.textarea.dispatchEvent(new Event("input"));
      await nextFrame();

      expect(getSignatureName(host)).toBe("CONTAR.SI(");
      expect(getActiveParamText(host)).toBe("range");

      const inSecondArg = view.textarea.value.indexOf('">0"') + 2;
      view.textarea.setSelectionRange(inSecondArg, inSecondArg);
      view.textarea.dispatchEvent(new Event("select"));
      await nextFrame();

      expect(getActiveParamText(host)).toBe("criteria");

      host.remove();
    } finally {
      document.documentElement.lang = prevLang;
    }
  });

  it("updates the active parameter as the cursor moves across commas", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=IF(A1>0,1,2)";

    const inFirstArg = view.textarea.value.indexOf(">") + 1;
    view.textarea.setSelectionRange(inFirstArg, inFirstArg);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();
    expect(getSignatureName(host)).toBe("IF(");
    expect(getActiveParamText(host)).toBe("logical_test");

    const inSecondArg = view.textarea.value.indexOf(",1") + 1;
    view.textarea.setSelectionRange(inSecondArg, inSecondArg);
    view.textarea.dispatchEvent(new Event("select"));
    await nextFrame();
    expect(getActiveParamText(host)).toBe("value_if_true");

    const inThirdArg = view.textarea.value.lastIndexOf(",2") + 1;
    view.textarea.setSelectionRange(inThirdArg, inThirdArg);
    view.textarea.dispatchEvent(new Event("select"));
    await nextFrame();
    expect(getActiveParamText(host)).toBe("[value_if_false]");

    host.remove();
  });

  it("uses the innermost function context for nested calls", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=IF(SUM(A1:A3)>0,1,2)";

    const inSumArg = view.textarea.value.indexOf("A1") + 1;
    view.textarea.setSelectionRange(inSumArg, inSumArg);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();
    expect(getSignatureName(host)).toBe("SUM(");
    expect(getActiveParamText(host)).toBe("number1");

    // Move the cursor back into the IF logical_test (after the SUM call closes).
    const inIfFirstArgAfterSum = view.textarea.value.indexOf(">") + 1;
    view.textarea.setSelectionRange(inIfFirstArgAfterSum, inIfFirstArgAfterSum);
    view.textarea.dispatchEvent(new Event("select"));
    await nextFrame();
    expect(getSignatureName(host)).toBe("IF(");
    expect(getActiveParamText(host)).toBe("logical_test");

    // Move into IF's second argument.
    const inIfSecondArg = view.textarea.value.indexOf(",1") + 1;
    view.textarea.setSelectionRange(inIfSecondArg, inIfSecondArg);
    view.textarea.dispatchEvent(new Event("select"));
    await nextFrame();
    expect(getActiveParamText(host)).toBe("value_if_true");

    host.remove();
  });

  it("keeps showing the innermost function hint when the cursor is after a closing paren (even with trailing whitespace)", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=ROUND(1, 2)\n";
    // Cursor after the closing paren.
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));
    await nextFrame();

    expect(getSignatureName(host)).toBe("ROUND(");
    // When positioned after the closing paren, treat the last argument as active.
    expect(getActiveParamText(host)).toBe("num_digits");

    host.remove();
  });
});
