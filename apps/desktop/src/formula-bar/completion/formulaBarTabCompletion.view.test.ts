/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { FormulaBarView } from "../FormulaBarView.js";
import { FormulaBarTabCompletionController } from "../../ai/completion/formulaBarTabCompletion.js";
import { getLocale, setLocale } from "../../i18n/index.js";

describe("FormulaBarView tab completion (integration)", () => {
  it("aborts in-flight backend completions on destroy()", async () => {
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    let observedSignal: AbortSignal | undefined;
    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      completionClient: {
        completeTabCompletion: vi.fn((req: any) => {
          observedSignal = req?.signal as AbortSignal | undefined;
          return new Promise((_resolve, reject) => {
            const signal = req?.signal as AbortSignal | undefined;
            if (!signal) {
              reject(new Error("missing abort signal"));
              return;
            }
            if (signal.aborted) {
              reject(signal.reason ?? new Error("aborted"));
              return;
            }
            signal.addEventListener(
              "abort",
              () => {
                reject(signal.reason ?? new Error("aborted"));
              },
              { once: true },
            );
          });
        }),
      },
    });

    try {
      view.focus({ cursor: "end" });
      view.textarea.value = "=1+";
      view.textarea.setSelectionRange(3, 3);
      view.textarea.dispatchEvent(new Event("input"));

      // Wait for the backend request to be initiated (the client sees the signal).
      for (let i = 0; i < 10 && !observedSignal; i += 1) {
        // eslint-disable-next-line no-await-in-loop
        await Promise.resolve();
      }

      expect(observedSignal).toBeDefined();
      expect(observedSignal?.aborted).toBe(false);

      completion.destroy();
      expect(observedSignal?.aborted).toBe(true);
      // Ensure no suggestion is applied after teardown.
      expect(view.model.aiSuggestion()).toBeNull();
    } finally {
      completion.destroy();
      host.remove();
    }
  });

  it("caps preview evaluation cell reads (MAX_CELL_READS) for large formulas", async () => {
    const doc = new DocumentController();
    // Preview evaluation intentionally avoids materializing sheets in an empty workbook.
    // Seed a single cell so Sheet1 exists and the evaluator will attempt to read the range.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    const getCellSpy = vi.spyOn(doc as any, "peekCell");

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      completionClient: {
        // Returning an insertion (not a full "=" formula) ensures we bypass the
        // rule-based starter-function suggestions for the bare "=" case.
        completeTabCompletion: async () => "SUM(A1:A6000)",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM(A1:A6000)");
    expect(view.model.aiSuggestionPreview()).toBe("(preview unavailable)");

    // Preview evaluation should be bounded (no unbounded cell reads for huge ranges).
    // (Implementations may short-circuit without reading any cells.)
    expect(getCellSpy.mock.calls.length).toBeLessThanOrEqual(5_000);

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.querySelector(".formula-bar-preview")?.textContent).toContain("(preview unavailable)");

    completion.destroy();
    host.remove();
  });

  it("uses Cursor backend completion for formula-body suggestions", async () => {
    const doc = new DocumentController();

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    let calls = 0;
    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      completionClient: {
        completeTabCompletion: async (_req) => {
          calls += 1;
          return "2";
        },
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(calls).toBe(1);
    expect(view.model.aiSuggestion()).toBe("=1+2");
    expect(view.model.aiGhostText()).toBe("2");
    expect(view.model.aiSuggestionPreview()).toBe(3);

    completion.destroy();
    host.remove();
  });

  it('includes workbook metadata in preview evaluation for CELL("filename")', async () => {
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
      getWorkbookFileMetadata: () => ({ directory: "/tmp/", filename: "Book.xlsx" }),
      completionClient: {
        completeTabCompletion: async () => 'CELL("filename")',
      },
    });

    view.focus({ cursor: "end" });
    // Use concatenation to ensure the preview evaluator treats the result as a string.
    view.textarea.value = '=""&';
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe('=""&CELL("filename")');
    expect(view.model.aiSuggestionPreview()).toBe("/tmp/[Book.xlsx]Sheet1");

    completion.destroy();
    host.remove();
  });

  it("evaluates preview values for implicit this-row structured references ([@Column])", async () => {
    const doc = new DocumentController();
    // Table range includes header row at A1:C1 and data rows at A2:C4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Total Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, "Calc");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 100);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 200);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 300);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Edit a cell inside the table (row 3 / col C).
    view.setActiveCell({ address: "C3", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getSheetNames: () => ["Sheet1"],
        getNamedRanges: () => [],
        getTables: () => [
          {
            name: "TableThisRow",
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 3,
            endCol: 2,
            columns: ["Amount", "Total Amount", "Calc"],
          },
        ],
        getCacheKey: () => "schema:v1",
      },
      completionClient: {
        completeTabCompletion: async () => "[@Amount]",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+[@Amount]");
    expect(view.model.aiSuggestionPreview()).toBe(21);

    completion.destroy();
    host.remove();
  });

  it("evaluates preview values for multi-column #This Row structured references (contiguous columns only)", async () => {
    const doc = new DocumentController();
    // Table range includes header row at A1:C1 and data rows at A2:C4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Total Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, "Calc");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 100);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 200);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 300);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Edit a cell inside the table (row 3 / col C).
    view.setActiveCell({ address: "C3", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getSheetNames: () => ["Sheet1"],
        getNamedRanges: () => [],
        getTables: () => [
          {
            name: "TableThisRow",
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 3,
            endCol: 2,
            columns: ["Amount", "Total Amount", "Calc"],
          },
        ],
        getCacheKey: () => "schema:v1",
      },
      completionClient: {
        completeTabCompletion: async () => "SUM([[#This Row],[Amount],[Total Amount]])",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM([[#This Row],[Amount],[Total Amount]])");
    expect(view.model.aiSuggestionPreview()).toBe(221);

    completion.destroy();
    host.remove();
  });

  it("evaluates preview values for implicit non-this-row structured references (e.g. [[#All],[Amount]])", async () => {
    const doc = new DocumentController();
    // Table range includes header row at A1:C1 and data rows at A2:C4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Total Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, "Calc");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 100);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 200);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 300);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Edit a cell inside the table (row 3 / col C).
    view.setActiveCell({ address: "C3", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getSheetNames: () => ["Sheet1"],
        getNamedRanges: () => [],
        getTables: () => [
          {
            name: "TableThisRow",
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 3,
            endCol: 2,
            columns: ["Amount", "Total Amount", "Calc"],
          },
        ],
        getCacheKey: () => "schema:v1",
      },
      completionClient: {
        completeTabCompletion: async () => "SUM([[#All],[Amount]])",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM([[#All],[Amount]])");
    expect(view.model.aiSuggestionPreview()).toBe(61);

    completion.destroy();
    host.remove();
  });

  it("evaluates preview values for whole-row implicit this-row structured references ([@])", async () => {
    const doc = new DocumentController();
    // Table range includes header row at A1:C1 and data rows at A2:C4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Total Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, "Calc");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 100);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 200);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 300);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Edit a cell inside the table (row 3 / col C).
    view.setActiveCell({ address: "C3", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getSheetNames: () => ["Sheet1"],
        getNamedRanges: () => [],
        getTables: () => [
          {
            name: "TableThisRow",
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 3,
            endCol: 2,
            columns: ["Amount", "Total Amount", "Calc"],
          },
        ],
        getCacheKey: () => "schema:v1",
      },
      completionClient: {
        completeTabCompletion: async () => "SUM([@])",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM([@])");
    // Current row is 20 + 200 (+ empty Calc column) = 220; plus leading 1 = 221.
    expect(view.model.aiSuggestionPreview()).toBe(221);

    completion.destroy();
    host.remove();
  });

  it("evaluates preview values for bracketed this-row column-range shorthand ([@[Col1]:[Col3]])", async () => {
    const doc = new DocumentController();
    // Table range includes header row at A1:C1 and data rows at A2:C4.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 1 }, "Total Amount");
    doc.setCellValue("Sheet1", { row: 0, col: 2 }, "Calc");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 1, col: 1 }, 100);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    doc.setCellValue("Sheet1", { row: 2, col: 1 }, 200);
    doc.setCellValue("Sheet1", { row: 3, col: 0 }, 30);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 300);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Edit a cell inside the table (row 3 / col C).
    view.setActiveCell({ address: "C3", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      schemaProvider: {
        getSheetNames: () => ["Sheet1"],
        getNamedRanges: () => [],
        getTables: () => [
          {
            name: "TableThisRow",
            sheetName: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 3,
            endCol: 2,
            columns: ["Amount", "Total Amount", "Calc"],
          },
        ],
        getCacheKey: () => "schema:v1",
      },
      completionClient: {
        completeTabCompletion: async () => "SUM([@[Amount]:[Total Amount]])",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM([@[Amount]:[Total Amount]])");
    expect(view.model.aiSuggestionPreview()).toBe(221);

    completion.destroy();
    host.remove();
  });

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

  it("uses the WASM partial parser (when available) to understand localized function names", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const view = new FormulaBarView(host, { onCommit: () => {} });
    const calls: unknown[][] = [];
    const engineStub = {
      parseFormulaPartial: async (...args: unknown[]) => {
        calls.push(args);
        // In de-DE, `COUNTIF` is localized as `ZÄHLENWENN`. The WASM parser returns the canonical
        // function name (`COUNTIF`) so completion can still use `FunctionRegistry` metadata even
        // when the user types a localized name (which does not exist in the canonical function catalog).
        return { context: { function: { name: "COUNTIF", argIndex: 0 } }, error: null, ast: {} };
      },
    };
    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      getEngineClient: () => engineStub as any,
    });
    try {
      for (let row = 0; row < 10; row += 1) {
        doc.setCellValue("Sheet1", { row, col: 0 }, row + 1);
      }
      view.setActiveCell({ address: "B11", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=ZÄHLENWENN(A";
      view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(calls).toHaveLength(1);
      expect(view.model.aiSuggestion()).toBe("=ZÄHLENWENN(A1:A10");
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
  });

  it("uses the WASM partial parser (when available) to infer function context for non-ASCII function names", async () => {
    const prevLocale = getLocale();
    setLocale("ar");

    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("Sheet1", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B11", input: "", value: null });

    const calls: unknown[][] = [];
    const engineStub = {
      parseFormulaPartial: async (...args: unknown[]) => {
        calls.push(args);
        // In some locales, function names may be localized/non-ASCII and won't match the canonical
        // function catalog. The WASM engine returns canonical function metadata so the completion
        // engine can still suggest appropriate argument values/ranges (e.g. range args for SUM).
        return { context: { function: { name: "SUM", argIndex: 0 } }, error: null, ast: {} };
      },
    };

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      getEngineClient: () => engineStub as any,
    });

    try {
      view.focus({ cursor: "end" });
      // Use an arbitrary non-ASCII identifier to simulate a localized function name.
      view.textarea.value = "=سوم(A";
      view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(calls).toHaveLength(1);
      expect(view.model.aiSuggestion()).toBe("=سوم(A1:A10)");
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
  });

  it("prefers localized function-name completion in de-DE (SU -> SUMME)", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
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

    try {
      view.focus({ cursor: "end" });
      view.textarea.value = "=SU";
      view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(view.model.aiSuggestion()).toBe("=SUMME(");
      expect(view.model.aiGhostText()).toBe("MME(");
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
  });

  it("previews localized range suggestions in de-DE (SUMME)", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
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

    try {
      view.focus({ cursor: "end" });
      view.textarea.value = "=SUMME(A";
      view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(view.model.aiSuggestion()).toBe("=SUMME(A1:A10)");
      expect(view.model.aiGhostText()).toBe("1:A10)");
      expect(view.model.aiSuggestionPreview()).toBe(55);
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
  });

  it("suggests localized starter functions for bare '=' in de-DE", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
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

    try {
      view.focus({ cursor: "end" });
      view.textarea.value = "=";
      view.textarea.setSelectionRange(1, 1);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(view.model.aiSuggestion()).toBe("=SUMME(");
      expect(view.model.aiGhostText()).toBe("SUMME(");
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
  });

  it("respects FormulaBarView tooling locale overrides when document.documentElement.lang differs", async () => {
    const prevLang = document.documentElement.lang;
    const prevDir = document.documentElement.dir;
    document.documentElement.lang = "en-US";
    document.documentElement.dir = "ltr";

    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} }, { getLocaleId: () => "de-DE" });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    try {
      view.focus({ cursor: "end" });
      view.textarea.value = "=";
      view.textarea.setSelectionRange(1, 1);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      expect(view.model.aiSuggestion()).toBe("=SUMME(");
      expect(view.model.aiGhostText()).toBe("SUMME(");
    } finally {
      completion.destroy();
      host.remove();
      document.documentElement.lang = prevLang;
      document.documentElement.dir = prevDir;
    }
  });
  
  it("falls back to the JS parser when the WASM partial parser throws", async () => {
    const prevLocale = getLocale();
    setLocale("de-DE");
    const doc = new DocumentController();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const view = new FormulaBarView(host, { onCommit: () => {} });
    const engineStub = {
      parseFormulaPartial: () => {
        throw new Error("engine not ready");
      },
    };
    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      getEngineClient: () => engineStub as any,
    });
    try {
      for (let row = 0; row < 10; row += 1) {
        doc.setCellValue("Sheet1", { row, col: 0 }, row + 1);
      }
      view.setActiveCell({ address: "B11", input: "", value: null });

      view.focus({ cursor: "end" });
      view.textarea.value = "=SUM(A";
      view.textarea.setSelectionRange(6, 6);
      view.textarea.dispatchEvent(new Event("input"));

      await completion.flushTabCompletion();

      // JS fallback still provides range suggestions.
      expect(view.model.aiSuggestion()).toBe("=SUM(A1:A10)");
    } finally {
      completion.destroy();
      host.remove();
      setLocale(prevLocale);
    }
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

  it("preserves the typed prefix case for named range suggestions", async () => {
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
    view.textarea.value = "=SUM(sal";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(salesData)");
    expect(view.model.aiGhostText()).toBe("esData)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    completion.destroy();
    host.remove();
  });

  it("suggests structured references and previews table column ranges", async () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
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
        getNamedRanges: () => [],
        getSheetNames: () => ["Sheet1"],
        getTables: () => [
          { name: "Table1", columns: ["Amount"], sheetName: "Sheet1", startRow: 0, startCol: 0, endRow: 2, endCol: 0 },
        ],
        getCacheKey: () => "tables:Table1",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(tab";
    view.textarea.setSelectionRange(8, 8);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM(table1[Amount])");
    expect(view.model.aiGhostText()).toBe("le1[Amount])");
    expect(view.model.aiSuggestionPreview()).toBe(30);

    completion.destroy();
    host.remove();
  });

  it("previews structured references with #All + whitespace", async () => {
    const doc = new DocumentController();
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "Amount");
    doc.setCellValue("Sheet1", { row: 1, col: 0 }, 10);
    doc.setCellValue("Sheet1", { row: 2, col: 0 }, 20);
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      completionClient: {
        completeTabCompletion: async () => "SUM(Table1[[ #All ], [Amount]])",
      },
      schemaProvider: {
        getNamedRanges: () => [],
        getSheetNames: () => ["Sheet1"],
        getTables: () => [
          { name: "Table1", columns: ["Amount"], sheetName: "Sheet1", startRow: 0, startCol: 0, endRow: 2, endCol: 0 },
        ],
        getCacheKey: () => "tables:Table1",
      },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=1+SUM(Table1[[ #All ], [Amount]])");
    expect(view.model.aiSuggestionPreview()).toBe(31);

    completion.destroy();
    host.remove();
  });

  it("does not create phantom sheets when suggesting sheet-qualified ranges", async () => {
    const doc = new DocumentController();
    // The completion controller may read from the active sheet as part of suggestion/preview
    // generation. Seed the active sheet so the test focuses on ensuring we *don't* materialize
    // unknown sheets (e.g. Sheet2) via read paths.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, 1);

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    // Use a row below the header so the range suggester scans upward (and would
    // previously materialize unknown sheets via `document.getCell()` reads).
    view.setActiveCell({ address: "B11", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      limits: { maxRows: 10_000, maxCols: 10_000 },
      // Provide a sheet list that includes a sheet that doesn't exist in the DocumentController yet.
      schemaProvider: {
        getNamedRanges: () => [],
        getTables: () => [],
        getSheetNames: () => ["Sheet2"],
        getCacheKey: () => "sheets:Sheet2",
      },
    });

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Sheet2!A";
    view.textarea.setSelectionRange(13, 13);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    // Completion may read from the active sheet for range inference, but should not
    // materialize the suggested (non-existent) Sheet2 via reads.
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
    expect(view.model.aiSuggestion()).toBe("=SUM(Sheet2!A:A)");

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

  it("suggests sheet-qualified ranges for sheet names requiring quotes", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      doc.setCellValue("sheet2", { row, col: 0 }, row + 1);
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "B11", input: "", value: null });

    const completion = new FormulaBarTabCompletionController({
      formulaBar: view,
      document: doc,
      getSheetId: () => "Sheet1",
      sheetNameResolver: {
        getSheetNameById: (id) => {
          if (id === "sheet2") return "My Sheet";
          if (id === "Sheet1") return "Sheet1";
          return null;
        },
        getSheetIdByName: (name) => {
          const key = name.trim().toLowerCase();
          if (key === "my sheet") return "sheet2";
          if (key === "sheet1") return "Sheet1";
          return null;
        },
      },
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM('My Sheet'!A";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(view.model.aiSuggestion()).toBe("=SUM('My Sheet'!A1:A10)");
    expect(view.model.aiGhostText()).toBe("1:A10)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=SUM('My Sheet'!A1:A10)");

    completion.destroy();
    host.remove();
  });

  it("uses sheet display names (not ids) after sheet rename", async () => {
    const doc = new DocumentController();
    for (let row = 0; row < 10; row += 1) {
      // Sheet id is still "Sheet2", but the display name is "Budget".
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
      sheetNameResolver: {
        getSheetNameById: (id) => {
          if (id === "Sheet2") return "Budget";
          if (id === "Sheet1") return "Sheet1";
          return null;
        },
        getSheetIdByName: (name) => {
          const key = name.trim().toLowerCase();
          if (key === "budget") return "Sheet2";
          if (key === "sheet1") return "Sheet1";
          return null;
        },
      },
      limits: { maxRows: 10_000, maxCols: 10_000 },
    });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(Budget!A";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    await completion.flushTabCompletion();

    expect(doc.getSheetIds()).not.toContain("Budget");
    expect(view.model.aiSuggestion()).toBe("=SUM(Budget!A1:A10)");
    expect(view.model.aiGhostText()).toBe("1:A10)");
    expect(view.model.aiSuggestionPreview()).toBe(55);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=SUM(Budget!A1:A10)");

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

  it("does not render a dangling summary separator when the signature has no summary", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=ABS(";
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("input"));

    // FormulaBarView renders on the next animation frame to keep long-formula edits responsive.
    await new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(() => resolve(), 0);
      }
    });

    const hint = host.querySelector<HTMLElement>('[data-testid="formula-hint"]');
    const signature = hint?.querySelector<HTMLElement>(".formula-bar-hint-signature");
    expect(signature?.textContent).toContain("ABS(number1)");
    expect(signature?.querySelector(".formula-bar-hint-token--paramActive")?.textContent).toBe("number1");
    expect(hint?.querySelector(".formula-bar-hint-summary-separator")).toBeNull();
    expect(hint?.querySelector(".formula-bar-hint-summary")).toBeNull();

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
