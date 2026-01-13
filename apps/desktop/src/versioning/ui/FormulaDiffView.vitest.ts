import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { FormulaDiffView } from "./FormulaDiffView.js";

describe("FormulaDiffView", () => {
  it("renders an explicit empty marker when both formulas are empty", () => {
    const html = renderToStaticMarkup(React.createElement(FormulaDiffView, { before: null, after: "   " }));
    expect(html).toContain("formula-diff-view--empty");
    expect(html).toContain("∅");
  });

  it("renders token operations with insert/delete classes", () => {
    const html = renderToStaticMarkup(React.createElement(FormulaDiffView, { before: "=A1", after: "=A2" }));
    expect(html).toContain("formula-diff-op--delete");
    expect(html).toContain("formula-diff-op--insert");
    expect(html).toContain("A1");
    expect(html).toContain("A2");
  });

  it("renders string tokens with quotes", () => {
    const html = renderToStaticMarkup(
      React.createElement(FormulaDiffView, { before: '=IF(A1="x",1,0)', after: '=IF(A1="y",1,0)' })
    );
    expect(html).toMatch(/(&quot;x&quot;|"x")/);
    expect(html).toMatch(/(&quot;y&quot;|"y")/);
  });
});
