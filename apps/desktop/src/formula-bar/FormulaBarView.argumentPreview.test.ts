/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

async function flushPreview(): Promise<void> {
  const flushRender = async (): Promise<void> => {
    // FormulaBarView coalesces renders via requestAnimationFrame when available.
    await new Promise<void>((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve());
      } else {
        setTimeout(resolve, 0);
      }
    });
  };

  // 1) Flush the render scheduled by input/selection changes.
  await flushRender();
  // 2) Allow the preview evaluation timer (setTimeout(..., 0)) to run.
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  // 3) Flush any promise microtasks from the preview provider + Promise.race.
  await Promise.resolve();
  await Promise.resolve();
  // 4) Flush the render scheduled after the preview resolves.
  await flushRender();
}

describe("FormulaBarView argument preview (integration)", () => {
  it("renders an evaluated preview for the active argument and updates as the cursor moves", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const provider = vi.fn((expr: string) => {
      if (expr === "A1") return true;
      if (expr === "B1") return 123;
      if (expr === "C1") return "#REF!";
      return "(preview unavailable)";
    });
    view.setArgumentPreviewProvider(provider);

    const formula = "=IF(A1, B1, C1)";
    view.textarea.value = formula;

    // Cursor inside first argument (A1).
    const cursorA1 = formula.indexOf("A1") + 1;
    view.textarea.setSelectionRange(cursorA1, cursorA1);
    view.textarea.dispatchEvent(new Event("input"));

    await flushPreview();

    const preview1 = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview1?.dataset.argStart).toBe(String(formula.indexOf("A1")));
    expect(preview1?.dataset.argEnd).toBe(String(formula.indexOf("A1") + 2));
    expect(preview1?.textContent).toBe("↳ A1  →  TRUE");

    // Cursor inside second argument (B1).
    const cursorB1 = formula.indexOf("B1") + 1;
    view.textarea.setSelectionRange(cursorB1, cursorB1);
    view.textarea.dispatchEvent(new Event("select"));

    await flushPreview();

    const preview2 = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview2?.dataset.argStart).toBe(String(formula.indexOf("B1")));
    expect(preview2?.dataset.argEnd).toBe(String(formula.indexOf("B1") + 2));
    expect(preview2?.textContent).toBe("↳ B1  →  123");

    // Cursor inside third argument (C1).
    const cursorC1 = formula.indexOf("C1") + 1;
    view.textarea.setSelectionRange(cursorC1, cursorC1);
    view.textarea.dispatchEvent(new Event("select"));

    await flushPreview();

    const preview3 = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview3?.dataset.argStart).toBe(String(formula.indexOf("C1")));
    expect(preview3?.dataset.argEnd).toBe(String(formula.indexOf("C1") + 2));
    expect(preview3?.textContent).toBe("↳ C1  →  #REF!");

    // Provider called with the active argument expression.
    expect(provider).toHaveBeenCalledWith("A1");
    expect(provider).toHaveBeenCalledWith("B1");
    expect(provider).toHaveBeenCalledWith("C1");

    host.remove();
  });

  it("treats escaped brackets inside structured refs as part of the argument expression", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const provider = vi.fn((expr: string) => {
      if (expr === "Table1[Total]],USD]") return 42;
      return "(preview unavailable)";
    });
    view.setArgumentPreviewProvider(provider);

    const formula = "=SUM(Table1[Total]],USD], 1)";
    view.textarea.value = formula;

    // Cursor inside the structured reference argument.
    const cursor = formula.indexOf("USD") + 1;
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("input"));

    await flushPreview();

    const preview = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview?.dataset.argStart).toBe(String(formula.indexOf("Table1")));
    expect(preview?.dataset.argEnd).toBe(String(formula.indexOf("Table1[Total]],USD]") + "Table1[Total]],USD]".length));
    expect(preview?.textContent).toBe("↳ Table1[Total]],USD]  →  42");
    expect(provider).toHaveBeenCalledWith("Table1[Total]],USD]");

    host.remove();
  });

  it("ignores commas inside quoted sheet names (sheet-qualified refs)", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    const provider = vi.fn((expr: string) => {
      if (expr === "'Budget,2025)'!A1") return 7;
      return "(preview unavailable)";
    });
    view.setArgumentPreviewProvider(provider);

    const formula = "=SUM('Budget,2025)'!A1, 1)";
    view.textarea.value = formula;

    // Cursor inside the sheet-qualified reference argument.
    const cursor = formula.indexOf("A1") + 1;
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("input"));

    await flushPreview();

    const preview = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview?.textContent).toBe("↳ 'Budget,2025)'!A1  →  7");
    expect(provider).toHaveBeenCalledWith("'Budget,2025)'!A1");

    host.remove();
  });

  it("collapses whitespace in the displayed argument expression", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.setArgumentPreviewProvider(() => true);

    const formula = "=IF(\n  SUM(A1:A2)\n    > 10,\n  B1,\n  C1\n)";
    view.textarea.value = formula;

    // Cursor inside the first IF argument but outside the nested SUM parentheses.
    const cursor = formula.indexOf("> 10") + 1;
    view.textarea.setSelectionRange(cursor, cursor);
    view.textarea.dispatchEvent(new Event("input"));

    await flushPreview();

    const preview = host.querySelector<HTMLElement>('[data-testid="formula-hint-arg-preview"]');
    expect(preview?.textContent).toBe("↳ SUM(A1:A2) > 10  →  TRUE");

    host.remove();
  });
});
