/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

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

  it("shows 1-letter function suggestions inside argument lists when they exist (=SUM(T → includes T)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(T";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("T");

    host.remove();
  });

  it("preserves typed casing when inserting (e.g. =vlo → =vlookup()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=vlo";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("VLOOKUP");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=vlookup(");

    host.remove();
  });

  it("preserves title-style casing when inserting (e.g. =Vlo → =Vlookup()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=Vlo";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=Vlookup(");

    host.remove();
  });

  it("supports _xlfn. prefix completion (=_xlfn.VLO → =_xlfn.VLOOKUP()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=_xlfn.VLO";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);
    expect(dropdown?.textContent).toContain("VLOOKUP");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=_xlfn.VLOOKUP(");

    host.remove();
  });

  it("preserves title-style casing after _xlfn. prefix (=_xlfn.Vlo → =_xlfn.Vlookup()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=_xlfn.Vlo";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=_xlfn.Vlookup(");

    host.remove();
  });

  it("supports Arrow navigation + Tab to accept (=VLOOKUP()", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true }));
    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.textarea.value).toBe("=VLOOKUP(");
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length);

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("accepts with '(' like Excel (commit character)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "(", cancelable: true, shiftKey: true }));

    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length);

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("accepts the clicked item", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    // Use a prefix with multiple options so we verify the click chooses the right one
    // (not just the current selection).
    view.textarea.value = "=COU";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const item = host.querySelector<HTMLButtonElement>(
      '[data-testid="formula-function-autocomplete-item"][data-name="COUNTIF"]',
    );
    expect(item).not.toBeNull();

    item?.click();

    expect(view.model.draft).toBe("=COUNTIF(");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length);

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("uses aria-activedescendant on the textarea while navigating", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    // Use a prefix with multiple matches so ArrowDown can advance selection.
    view.textarea.value = "=COU";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    // Initial selection should populate aria-activedescendant.
    const initial = view.textarea.getAttribute("aria-activedescendant");
    expect(typeof initial).toBe("string");
    expect(initial?.length).toBeGreaterThan(0);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true }));
    const afterDown = view.textarea.getAttribute("aria-activedescendant");
    expect(afterDown).not.toBe(initial);

    // Closing clears aria-activedescendant.
    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));
    expect(view.textarea.hasAttribute("aria-activedescendant")).toBe(false);

    host.remove();
  });

  it("uses Shift+Tab to commit (and does not accept) when the dropdown is open", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    const e = new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=VLO", { reason: "tab", shift: true });
    expect(view.model.isEditing).toBe(false);
    // Should not accept completion when Shift+Tab is used for commit/navigation semantics.
    expect(view.model.activeCell.input).toBe("=VLO");
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

  it("accepts with Tab first, then uses a second Tab to commit (dropdown should not steal commit forever)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    const first = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(first);

    expect(first.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    const second = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(second);

    expect(second.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=VLOOKUP(", { reason: "tab", shift: false });
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("accepts with Enter first, then uses a second Enter to commit the edit", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    const first = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(first);

    expect(first.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    const second = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(second);

    expect(second.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=VLOOKUP(", { reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(view.model.activeCell.input).toBe("=VLOOKUP(");

    host.remove();
  });

  it("prefers dropdown completion over AI ghost text when open", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    // Configure a conflicting AI suggestion so we can observe precedence.
    view.setAiSuggestion("=VLOAI");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=VLOOKUP(");
    expect(view.model.aiSuggestion()).toBeNull();

    host.remove();
  });

  it("does not duplicate an existing opening paren (e.g. =VLO() → =VLOOKUP())", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO()";
    // Caret before the existing "(" so autocomplete still triggers.
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
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

  it("closes if the formula bar becomes read-only while editing", () => {
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
    expect(view.model.isEditing).toBe(true);

    view.setReadOnly(true);

    expect(view.model.isEditing).toBe(false);
    expect(dropdown?.hasAttribute("hidden")).toBe(true);

    host.remove();
  });

  it("closes during IME composition", () => {
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

    view.textarea.dispatchEvent(new Event("compositionstart"));
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
