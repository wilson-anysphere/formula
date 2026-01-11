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

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight?.textContent).toBe("=SUM(A1:A10)");
    expect(highlight?.querySelectorAll(".formula-bar-ghost")).toHaveLength(1);
    expect(highlight?.querySelector(".formula-bar-ghost")?.textContent).toBe("1:A10)");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=SUM(A1:A10)");

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
    expect(highlight?.textContent).toBe("=VLOOKUP(");
    expect(highlight?.querySelectorAll(".formula-bar-ghost")).toHaveLength(1);
    expect(highlight?.querySelector(".formula-bar-ghost")?.textContent).toBe("OKUP(");

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", cancelable: true }));
    expect(view.model.draft).toBe("=VLOOKUP(");

    completion.destroy();
    host.remove();
  });
});
