/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView named range highlights", () => {
  it("colors resolved named ranges while editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let lastHighlights: any[] = [];
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onReferenceHighlights: (highlights) => {
        lastHighlights = highlights;
      },
    });

    view.model.setNameResolver((name) => (name === "SalesData" ? { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } : null));
    view.setActiveCell({ address: "A1", input: "=SUM(SalesData)", value: null });
    view.focus({ cursor: "end" });

    // Move caret inside the named range token so it becomes active.
    const caret = "=SUM(".length + 1;
    view.textarea.setSelectionRange(caret, caret);
    view.textarea.dispatchEvent(new Event("select"));

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-ref-index="0"]');
    expect(refSpan?.textContent).toBe("SalesData");
    // Named ranges are tokenized as identifiers, but should still receive the per-reference styling.
    expect(refSpan?.dataset.kind).toBe("identifier");

    expect(lastHighlights).toHaveLength(1);
    expect(lastHighlights[0]?.text).toBe("SalesData");

    host.remove();
  });
});

