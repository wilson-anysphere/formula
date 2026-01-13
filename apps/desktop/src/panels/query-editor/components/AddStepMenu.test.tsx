// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DataTable, type Query, type QueryOperation } from "@formula/power-query";

import { setLocale } from "../../../i18n/index.js";
import { AddStepMenu } from "./AddStepMenu";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

async function flushMicrotasks(count = 5): Promise<void> {
  for (let i = 0; i < count; i++) await Promise.resolve();
}

function baseQuery(): Query {
  return { id: "q1", name: "Query 1", source: { type: "range", range: { values: [] } }, steps: [] };
}

function findButtonByText(host: HTMLElement, text: string): HTMLButtonElement {
  const buttons = Array.from(host.querySelectorAll("button"));
  const match = buttons.find((btn) => btn.textContent?.trim() === text);
  if (!match) {
    const available = buttons.map((btn) => btn.textContent?.trim()).filter(Boolean);
    throw new Error(`Could not find button '${text}'. Available: ${available.join(", ")}`);
  }
  return match as HTMLButtonElement;
}

describe("AddStepMenu", () => {
  let host: HTMLDivElement | null = null;
  let root: ReturnType<typeof createRoot> | null = null;

  beforeEach(() => {
    setLocale("en-US");
    host = document.createElement("div");
    document.body.appendChild(host);
    root = createRoot(host);
  });

  afterEach(async () => {
    if (root) {
      await act(async () => {
        root?.unmount();
      });
    }
    host?.remove();
    host = null;
    root = null;
    vi.restoreAllMocks();
  });

  it("renders expected categories + operations", async () => {
    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Sales", type: "number" },
      ],
      [],
    );

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} aiContext={{ query: baseQuery(), preview }} />);
    });

    const addStep = findButtonByText(host!, "+ Add step");
    await act(async () => {
      addStep.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(host?.textContent).toContain("Rows");
    expect(host?.textContent).toContain("Columns");
    expect(host?.textContent).toContain("Transform");

    for (const label of [
      "Filter Rows",
      "Sort",
      "Remove Columns",
      "Keep Columns",
      "Rename Columns",
      "Change Type",
      "Split Column",
      "Group By",
    ]) {
      expect(host?.textContent).toContain(label);
    }
  });

  it("constructs a minimally-valid operation for menu selections", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Filter Rows").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledTimes(1);
    expect(onAddStep).toHaveBeenCalledWith({
      type: "filterRows",
      predicate: { type: "comparison", column: "Region", operator: "isNotNull" },
    });
  });

  it("disables schema-dependent operations when preview schema is missing", async () => {
    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} aiContext={{ query: baseQuery(), preview: null }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const filterButton = findButtonByText(host!, "Filter Rows");
    expect(filterButton.disabled).toBe(true);
    expect(filterButton.title).toContain("Preview schema required");
    expect(host?.textContent).toContain("Preview schema required");
  });

  it("closes the operation menu on outside click and Escape", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(host!.querySelector(".query-editor-add-step__menu-popover")).toBeTruthy();

    await act(async () => {
      document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    });

    expect(host!.querySelector(".query-editor-add-step__menu-popover")).toBeNull();

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(host!.querySelector(".query-editor-add-step__menu-popover")).toBeTruthy();

    await act(async () => {
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    });

    expect(host!.querySelector(".query-editor-add-step__menu-popover")).toBeNull();
  });

  it("handles empty AI intent and renders returned suggestions with readable labels", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query = baseQuery();
    const aiContext = { query, preview };

    const deferred = (() => {
      let resolve!: (value: QueryOperation[]) => void;
      const promise = new Promise<QueryOperation[]>((res) => {
        resolve = res;
      });
      return { promise, resolve };
    })();

    const onAiSuggest = vi.fn(() => deferred.promise);

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} onAiSuggest={onAiSuggest} aiContext={aiContext} />);
    });

    const input = host!.querySelector("input.query-editor-add-step__ai-input") as HTMLInputElement;
    const suggestButton = host!.querySelector("button.query-editor-add-step__ai-button") as HTMLButtonElement;

    await act(async () => {
      input.value = "   ";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    expect(suggestButton.disabled).toBe(true);
    expect(onAiSuggest).not.toHaveBeenCalled();

    await act(async () => {
      input.value = "filter to non-null";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });

    expect(suggestButton.disabled).toBe(false);

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAiSuggest).toHaveBeenCalledWith("filter to non-null", aiContext);
    expect(suggestButton.textContent).toContain("Suggesting");

    await act(async () => {
      deferred.resolve([{ type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } }]);
      await flushMicrotasks(10);
    });

    const suggestionButtons = host!.querySelectorAll("button.query-editor-add-step__suggestion");
    expect(suggestionButtons.length).toBe(1);
    expect(suggestionButtons[0]?.textContent).toBe("Filter Rows (Region)");
  });

  it("clears existing AI suggestions when the intent changes or after applying one", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query = baseQuery();
    const aiContext = { query, preview };

    const onAiSuggest = vi.fn(async () => [
      { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } },
    ]);
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} onAiSuggest={onAiSuggest} aiContext={aiContext} />);
    });

    const input = host!.querySelector("input.query-editor-add-step__ai-input") as HTMLInputElement;
    const suggestButton = host!.querySelector("button.query-editor-add-step__ai-button") as HTMLButtonElement;

    await act(async () => {
      input.value = "filter";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(10);
    });

    expect(host!.querySelectorAll("button.query-editor-add-step__suggestion").length).toBe(1);

    await act(async () => {
      input.value = "filter again";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });

    expect(host!.querySelectorAll("button.query-editor-add-step__suggestion").length).toBe(0);

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(10);
    });

    const suggestionButton = host!.querySelector("button.query-editor-add-step__suggestion") as HTMLButtonElement;
    expect(suggestionButton).toBeTruthy();

    await act(async () => {
      suggestionButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledTimes(1);
    expect(host!.querySelectorAll("button.query-editor-add-step__suggestion").length).toBe(0);
  });

  it("shows an error message when AI suggestion fails", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query = baseQuery();
    const aiContext = { query, preview };
    const onAiSuggest = vi.fn(async () => {
      throw new Error("AI exploded");
    });

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} onAiSuggest={onAiSuggest} aiContext={aiContext} />);
    });

    const input = host!.querySelector("input.query-editor-add-step__ai-input") as HTMLInputElement;
    const suggestButton = host!.querySelector("button.query-editor-add-step__ai-button") as HTMLButtonElement;

    await act(async () => {
      input.value = "do something";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(5);
    });

    expect(host?.textContent).toContain("AI exploded");
    // Error state should not be confused with the empty suggestions state.
    expect(host?.textContent).not.toContain("No suggestions.");
  });
});
