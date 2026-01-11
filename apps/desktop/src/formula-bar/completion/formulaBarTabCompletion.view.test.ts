/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { FormulaBarView } from "../FormulaBarView.js";
import { FormulaBarTabCompletionController } from "../../ai/completion/formulaBarTabCompletion.js";

describe("FormulaBarView tab completion (integration)", () => {
  it("suggests contiguous ranges for SUM when typing a column reference", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("Sheet1", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B11", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(A";
    view.textarea.setSelectionRange(6, 6);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(A1:A10)");
    expect(view.model.aiGhostText()).toBe("1:A10)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toContain("=SUM(A1:A10)");
    expect(highlight?.querySelectorAll(".formula-bar-ghost")).toHaveLength(1);
    expect(highlight?.querySelector(".formula-bar-ghost")?.textContent).toBe("1:A10)");
    expect(highlight?.querySelector(".formula-bar-preview")?.textContent).toContain("55");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=SUM(A1:A10)");

    completion.destroy();
    host.remove();
  });

  it("suggests named ranges when typing a range argument", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("Sheet1", { row, col: 0 }, row + 1);
    }
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getNamedRanges: () => [{ name: "SalesData", range: "Sheet1!A1:A10" }],
        getTables: () => [],
        getSheetNames: () => ["Sheet1"],
        getCacheKey: () => "namedRanges:SalesData",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sal";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(SalesData)");
    expect(view.model.aiGhostText()).toBe("esData)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toContain("=SUM(SalesData)");
    expect(highlight?.querySelector(".formula-bar-preview")?.textContent).toContain("55");

    completion.destroy();
    host.remove();
  });

  it("previews named ranges that refer to another sheet", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("Sheet2", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getNamedRanges: () => [{ name: "SalesData", range: "Sheet2!A1:A10" }],
        getTables: () => [],
        getSheetNames: () => ["Sheet1", "Sheet2"],
        getCacheKey: () => "namedRanges:Sheet2",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sal";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(SalesData)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    completion.destroy();
    host.remove();
  });

  it("previews named ranges that refer to a sheet requiring quotes", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("My Sheet", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getNamedRanges: () => [{ name: "SalesData", range: "'My Sheet'!A1:A10" }],
        getTables: () => [],
        getSheetNames: () => ["Sheet1", "My Sheet"],
        getCacheKey: () => "namedRanges:My Sheet",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sal";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(SalesData)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    completion.destroy();
    host.remove();
  });

  it("previews named ranges that refer to a numeric sheet id (requires quotes)", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("2024", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getNamedRanges: () => [{ name: "SalesData", range: "'2024'!A1:A10" }],
        getTables: () => [],
        getSheetNames: () => ["Sheet1", "2024"],
        getCacheKey: () => "namedRanges:2024",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sal";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(SalesData)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    completion.destroy();
    host.remove();
  });

  it("treats formulas that evaluate to blank as non-empty when suggesting ranges", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      // Formula that evaluates to empty string, but should still count as "non-empty"
      // for used-range detection in tab completion.
      doc.setCellFormula("Sheet1", { row, col: 0 }, '=""');
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B11", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(A";
    view.textarea.setSelectionRange(6, 6);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(A1:A10)");
    expect(view.model.aiGhostText()).toBe("1:A10)");
    expect(view.model.aiSuggestionPreview()).toBe(0);

    completion.destroy();
    host.remove();
  });

  it("previews sheet-qualified range suggestions", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("Sheet2", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B11", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sheet2!A";
    view.textarea.setSelectionRange(13, 13);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(Sheet2!A1:A10)");
    expect(view.model.aiGhostText()).toBe("1:A10)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toContain("=SUM(Sheet2!A1:A10)");
    expect(highlight?.querySelector(".formula-bar-preview")?.textContent).toContain("55");

    completion.destroy();
    host.remove();
  });

  it("suggests function name completion (=VLO → VLOOKUP()", async () => {
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=VLOOKUP(");
    expect(view.model.aiGhostText()).toBe("OKUP(");

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toContain("=VLOOKUP(");
    expect(highlight?.querySelectorAll(".formula-bar-ghost")).toHaveLength(1);
    expect(highlight?.querySelector(".formula-bar-ghost")?.textContent).toBe("OKUP(");
    expect(highlight?.querySelector(".formula-bar-preview")?.textContent).toContain("(preview unavailable)");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=VLOOKUP(");

    completion.destroy();
    host.remove();
  });

  it("clears suggestions when the selection is not collapsed", async () => {
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=VLOOKUP(");
    expect(view.model.aiGhostText()).toBe("OKUP(");

    view.textarea.setSelectionRange(0, 2);
    view.textarea.dispatchEvent(new Event("select"));

    expect(view.model.aiSuggestion()).toBeNull();
    expect(view.model.aiGhostText()).toBe("");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=VLO");

    completion.destroy();
    host.remove();
  });

  it("does not render a dangling summary separator when the signature has no summary", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=ABS(";
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("input"));

    const hint = host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
    // Active parameter is rendered in brackets (see FormulaBarView paramActive formatting).
    expect(hint?.textContent).toContain("ABS([number1])");
    expect(hint?.textContent).not.toContain("—");

    host.remove();
  });

  it("ignores stale suggestions when the active sheet changes mid-request", async () => {
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    let sheetId = "Sheet1";
    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => sheetId,
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    // Simulate a sheet switch happening before the async completion resolves.
    sheetId = "Sheet2";

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBeNull();
    expect(view.model.aiGhostText()).toBe("");

    completion.destroy();
    host.remove();
  });
});
