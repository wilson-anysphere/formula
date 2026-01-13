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

describe("FormulaBarView structured reference highlights", () => {
  it("colors resolved structured table references while editing", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let lastHighlights: any[] = [];
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onReferenceHighlights: (highlights) => {
        lastHighlights = highlights;
      },
    });

    const tables = new Map([
      [
        "Table1",
        {
          name: "Table1",
          sheetName: "Sheet1",
          startRow: 0,
          startCol: 0,
          endRow: 3,
          endCol: 1,
          columns: ["Item", "Qty"],
        },
      ],
    ]);

    view.model.setExtractFormulaReferencesOptions({ tables });
    view.setActiveCell({ address: "A1", input: "=SUM(Table1[Qty])", value: null });
    view.focus({ cursor: "end" });

    // Move caret inside the structured ref token so it becomes active.
    const caret = "=SUM(".length + "Table1[".length + 1;
    view.textarea.setSelectionRange(caret, caret);
    view.textarea.dispatchEvent(new Event("select"));

    await nextFrame();

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpan = highlight?.querySelector<HTMLElement>('span[data-ref-index="0"]');
    expect(refSpan?.textContent).toBe("Table1[Qty]");
    expect(refSpan?.dataset.kind).toBe("reference");

    expect(lastHighlights).toHaveLength(1);
    expect(lastHighlights[0]?.text).toBe("Table1[Qty]");
    expect(lastHighlights[0]?.range).toEqual({ sheet: "Sheet1", startRow: 1, startCol: 1, endRow: 3, endCol: 1 });

    host.remove();
  });
});
