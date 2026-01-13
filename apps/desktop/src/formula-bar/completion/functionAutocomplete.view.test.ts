/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "../FormulaBarView.js";

describe("FormulaBarView function autocomplete dropdown", () => {
  it("shows a dropdown with matching functions (=VLO → includes VLOOKUP)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown).not.toBeNull();
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("VLOOKUP");

    host.remove();
  });

  it("shows function suggestions inside argument lists (=SUM(IF → includes IF)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(IF";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("IF");

    host.remove();
  });

  it("preserves typed casing when inserting (e.g. =vlo → =vlookup()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=vlo";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("VLOOKUP");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(view.model.draft).toBe("=vlookup(");

    host.remove();
  });

  it("supports _xlfn. prefix completion (=_xlfn.VLO → =_xlfn.VLOOKUP()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=_xlfn.VLO";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("VLOOKUP");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(view.model.draft).toBe("=_xlfn.VLOOKUP(");

    host.remove();
  });

  it("supports Arrow navigation + Tab to accept (=VLOOKUP()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true }));
    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(view.textarea.value).toBe("=VLOOKUP(");
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length);

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("accepts with Enter (and does not commit the edit)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let committed = false;
    const view = new FormulaBarView(host, {
      onCommit: () => {
        committed = true;
      },
    });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));

    expect(committed).toBe(false);
    expect(view.textarea.value).toBe("=VLOOKUP(");
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.model.isEditing).toBe(true);

    host.remove();
  });

  it("prefers dropdown completion over AI ghost text when open", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    // Configure a conflicting AI suggestion so we can observe precedence.
    view.setAiSuggestion("=VLOAI");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.model.aiSuggestion()).toBeNull();

    host.remove();
  });

  it("does not duplicate an existing opening paren (e.g. =VLO() → =VLOOKUP())", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO()";
    // Caret before the existing "(" so autocomplete still triggers.
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(view.model.draft).toBe("=VLOOKUP()");
    // Cursor inside the existing parens.
    expect(view.textarea.selectionStart).toBe(9);
    expect(view.textarea.selectionEnd).toBe(9);

    host.remove();
  });

  it("closes when starting a range selection (to avoid stealing Tab/Enter)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    view.beginRangeSelection({ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, "Sheet1");

    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("closes the dropdown on Escape", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));

    expect(dropdown?.hasAttribute("hidden")).toBe(true);
    // Escape should close the dropdown before cancelling the edit.
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=VLO");

    host.remove();
  });
});
