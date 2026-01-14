/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView highlight HTML escaping", () => {
  it("escapes reference text that looks like an HTML entity (e.g. &lt;) so highlight rendering is safe", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const formula = "=SUM('A&lt;B'!A1)";
    view.setActiveCell({ address: "A1", input: formula, value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toBe(formula);

    const refSpan = highlight?.querySelector<HTMLElement>('span[data-kind="reference"]');
    expect(refSpan?.textContent).toBe("'A&lt;B'!A1");

    host.remove();
  });
});

