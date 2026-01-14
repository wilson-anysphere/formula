/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView Alt+Enter auto-indentation", () => {
  it("inserts a newline + indentation based on paren nesting", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const commits: string[] = [];
    const view = new FormulaBarView(host, { onCommit: (text) => commits.push(text) });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const formula = `=IF(SUM(A1:A10)>0,\n"ok",\n"bad")`;
    view.textarea.value = formula;
    view.textarea.setSelectionRange(formula.length, formula.length);
    view.textarea.dispatchEvent(new Event("input"));

    // Insert a line break immediately after `=IF(` (before `SUM…`).
    const cursor = formula.indexOf("SUM");
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("select"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", altKey: true, cancelable: true }));

    const expected = `=IF(\n  SUM(A1:A10)>0,\n"ok",\n"bad")`;
    expect(view.textarea.value).toBe(expected);
    expect(view.model.draft).toBe(expected);
    expect(view.textarea.selectionStart).toBe(cursor + 3); // "\n" + 2 spaces
    expect(view.textarea.selectionEnd).toBe(cursor + 3);
    expect(commits).toEqual([]);

    host.remove();
  });

  it("indents deeper inside nested function calls", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const commits: string[] = [];
    const view = new FormulaBarView(host, { onCommit: (text) => commits.push(text) });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const formula = `=IF(SUM(A1:A10)>0,\n"ok",\n"bad")`;
    view.textarea.value = formula;
    view.textarea.setSelectionRange(formula.length, formula.length);
    view.textarea.dispatchEvent(new Event("input"));

    // Insert a line break immediately after `SUM(` (before `A1:A10`).
    const cursor = formula.indexOf("A1:A10");
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("select"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", altKey: true, cancelable: true }));

    const expected = `=IF(SUM(\n    A1:A10)>0,\n"ok",\n"bad")`;
    expect(view.textarea.value).toBe(expected);
    expect(view.model.draft).toBe(expected);
    expect(view.textarea.selectionStart).toBe(cursor + 5); // "\n" + 4 spaces
    expect(view.textarea.selectionEnd).toBe(cursor + 5);
    expect(commits).toEqual([]);

    host.remove();
  });

  it("does not regress Enter-to-commit behavior", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let committed: string | null = null;
    const view = new FormulaBarView(host, { onCommit: (text) => (committed = text) });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.textarea.value = "=SUM(A1:A10)";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));

    expect(committed).toBe("=SUM(A1:A10)");
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("ignores parentheses inside string literals", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const commits: string[] = [];
    const view = new FormulaBarView(host, { onCommit: (text) => commits.push(text) });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const formula = `=IF("(",1,2`;
    view.textarea.value = formula;
    view.textarea.setSelectionRange(formula.length, formula.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", altKey: true, cancelable: true }));

    // Should indent for the `IF(` call, but ignore the `(` inside the string literal.
    const expected = `=IF("(",1,2\n  `;
    expect(view.textarea.value).toBe(expected);
    expect(view.model.draft).toBe(expected);
    expect(view.textarea.selectionStart).toBe(expected.length);
    expect(view.textarea.selectionEnd).toBe(expected.length);
    expect(commits).toEqual([]);

    host.remove();
  });
});
