/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";
import { parseA1Range } from "../spreadsheet/a1.js";

describe("FormulaBarView hover previews", () => {
  it("emits a range for sheet-qualified references when hovering in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hovered = null as ReturnType<typeof parseA1Range>;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRange: (range) => {
        hovered = range;
      },
    });

    view.setActiveCell({ address: "A1", input: "=Sheet2!A1:B2", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("Sheet2!A1:B2");

    refSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hovered).toEqual(parseA1Range("A1:B2"));

    host.remove();
  });

  it("emits a range for named ranges when hovering in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hovered = null as ReturnType<typeof parseA1Range>;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRange: (range) => {
        hovered = range;
      },
    });

    view.model.setNameResolver((name) =>
      name === "SalesData" ? { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } : null
    );

    view.setActiveCell({ address: "A1", input: "=SUM(SalesData)", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const identifierSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="identifier"]') ?? []);
    const nameSpan = identifierSpans.find((s) => s.textContent === "SalesData");
    expect(nameSpan?.textContent).toBe("SalesData");

    nameSpan?.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hovered).toEqual(parseA1Range("A1:B2"));

    host.remove();
  });

  it("emits a range for sheet-qualified references while editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hovered = null as ReturnType<typeof parseA1Range>;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRange: (range) => {
        hovered = range;
      },
    });

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.textarea.value = "=Sheet2!A1:B2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    expect(hovered).toEqual(parseA1Range("A1:B2"));

    host.remove();
  });
});
