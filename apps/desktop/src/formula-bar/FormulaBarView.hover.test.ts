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

  it("emits a range for structured references while editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 0,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const formula = "=SUM(Table1[Amount])";
    view.textarea.value = formula;
    const caret = formula.indexOf("Amount") + 1;
    view.textarea.setSelectionRange(caret, caret);
    view.textarea.dispatchEvent(new Event("input"));

    expect(hoveredText).toBe("Table1[Amount]");
    expect(hoveredRange).toEqual(parseA1Range("A2:A3"));

    host.remove();
  });

  it("highlights structured references and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    // Table1 is A1:A3 with a header row at A1, so Table1[Amount] should resolve to A2:A3.
    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 0,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[Amount])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? [];
    expect(refSpans).toHaveLength(1);
    const refSpan = refSpans[0]!;
    expect(refSpan.textContent).toBe("Table1[Amount]");

    refSpan.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hoveredText).toBe("Table1[Amount]");
    expect(hoveredRange).toEqual(parseA1Range("A2:A3"));

    host.remove();
  });

  it("highlights nested structured refs (#All) and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 0,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[[#All],[Amount]])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? [];
    expect(refSpans).toHaveLength(1);
    const refSpan = refSpans[0]!;
    expect(refSpan.textContent).toBe("Table1[[#All],[Amount]]");

    refSpan.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hoveredText).toBe("Table1[[#All],[Amount]]");
    // #All includes the header row; with a 3-row table this is A1:A3.
    expect(hoveredRange).toEqual(parseA1Range("A1:A3"));

    host.remove();
  });

  it("highlights structured ref specifiers (Table1[#All]) and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount", "Other"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 1,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[#All])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? [];
    expect(refSpans).toHaveLength(1);
    const refSpan = refSpans[0]!;
    expect(refSpan.textContent).toBe("Table1[#All]");

    refSpan.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hoveredText).toBe("Table1[#All]");
    expect(hoveredRange).toEqual(parseA1Range("A1:B3"));

    host.remove();
  });

  it("highlights selector-qualified structured refs (Table1[[#Headers],[Amount]]) and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount", "Other"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 1,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[[#Headers],[Amount]])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? [];
    expect(refSpans).toHaveLength(1);
    const refSpan = refSpans[0]!;
    expect(refSpan.textContent).toBe("Table1[[#Headers],[Amount]]");

    refSpan.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hoveredText).toBe("Table1[[#Headers],[Amount]]");
    expect(hoveredRange).toEqual(parseA1Range("A1"));

    host.remove();
  });

  it("highlights selector-qualified structured refs (Table1[[#Data],[Amount]]) and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount", "Other"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 1,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[[#Data],[Amount]])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? [];
    expect(refSpans).toHaveLength(1);
    const refSpan = refSpans[0]!;
    expect(refSpan.textContent).toBe("Table1[[#Data],[Amount]]");

    refSpan.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));

    expect(hoveredText).toBe("Table1[[#Data],[Amount]]");
    // #Data excludes the header row; in a 3-row table (header + 2 data rows) that's A2:A3.
    expect(hoveredRange).toEqual(parseA1Range("A2:A3"));

    host.remove();
  });

  it("highlights structured ref specifiers (Table1[#Data]/[#Totals]) and emits hover previews (with text) in view mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let hoveredRange = null as ReturnType<typeof parseA1Range>;
    let hoveredText = null as string | null;
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onHoverRangeWithText: (range, refText) => {
        hoveredRange = range;
        hoveredText = refText;
      },
    });

    view.model.setExtractFormulaReferencesOptions({
      tables: [
        {
          name: "Table1",
          columns: ["Amount", "Other"],
          startRow: 0,
          startCol: 0,
          endRow: 2,
          endCol: 1,
          sheetName: "Sheet1",
        },
      ],
    });

    view.setActiveCell({ address: "A1", input: "=SUM(Table1[#Data],Table1[#Totals])", value: null });

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    const refSpans = Array.from(highlight?.querySelectorAll<HTMLElement>('span[data-kind="reference"]') ?? []);
    expect(refSpans.map((s) => s.textContent)).toEqual(["Table1[#Data]", "Table1[#Totals]"]);

    // Hover #Data -> excludes header row (A2:B3).
    refSpans[0]!.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    expect(hoveredText).toBe("Table1[#Data]");
    expect(hoveredRange).toEqual(parseA1Range("A2:B3"));

    // Hover #Totals -> last row only (A3:B3).
    refSpans[1]!.dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
    expect(hoveredText).toBe("Table1[#Totals]");
    expect(hoveredRange).toEqual(parseA1Range("A3:B3"));

    host.remove();
  });
});
