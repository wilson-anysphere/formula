/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView reference token click selection toggle", () => {
  it("selects a reference token on first click and toggles back to a caret on second click", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.textarea.value = "=A1+B1";

    // Place caret inside "A1" (between A and 1), then click: Excel UX should expand
    // selection to the full reference token.
    view.textarea.setSelectionRange(2, 2);
    view.textarea.dispatchEvent(new Event("input"));
    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(3);

    // Second click on the same token should toggle back to a caret for manual edits.
    //
    // In browsers, the selection typically collapses to a caret (and can emit a `select`
    // event) before the `click` handler runs. Emulate that ordering to ensure the toggle
    // logic is resilient even when `activeReferenceIndex` and `selectedReferenceIndex`
    // temporarily differ.
    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.setSelectionRange(2, 2);
    view.textarea.dispatchEvent(new Event("select"));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(2);
    expect(view.textarea.selectionEnd).toBe(2);

    // Repeat for the second reference to ensure the correct token is selected.
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(4);
    expect(view.textarea.selectionEnd).toBe(6);

    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("select"));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(5);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });

  it("supports the same click-to-select / click-again-to-edit toggle for named ranges", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    // Enable named-range extraction so identifier tokens can participate in the same UX.
    view.model.setNameResolver((name) => {
      if (name !== "Sales") return null;
      return { startRow: 0, startCol: 0, endRow: 9, endCol: 0 };
    });

    view.textarea.value = "=Sales+B1";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(6);

    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("select"));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(3);
    expect(view.textarea.selectionEnd).toBe(3);

    host.remove();
  });

  it("supports the same click-to-select / click-again-to-edit toggle for structured references", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    // Enable structured reference extraction via table metadata.
    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          sheet: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 10,
          endCol: 0,
        },
      ],
    });

    const refText = "Table1[Amount]";
    view.textarea.value = `=${refText}+B1`;

    const refStart = view.textarea.value.indexOf(refText);
    const refEnd = refStart + refText.length;
    const caret = refStart + 3;

    view.textarea.setSelectionRange(caret, caret);
    view.textarea.dispatchEvent(new Event("input"));
    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(refStart);
    expect(view.textarea.selectionEnd).toBe(refEnd);

    view.textarea.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    view.textarea.setSelectionRange(caret, caret);
    view.textarea.dispatchEvent(new Event("select"));
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(caret);
    expect(view.textarea.selectionEnd).toBe(caret);

    host.remove();
  });
});
