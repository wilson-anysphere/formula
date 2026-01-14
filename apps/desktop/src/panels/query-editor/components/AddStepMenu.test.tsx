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

function setTextInputValue(input: HTMLInputElement, value: string): void {
  // React tracks input values by patching the element's `value` property. When we
  // assign to `input.value` directly, React may treat the subsequent `input`
  // event as a no-op because the tracker has already been updated. Calling the
  // native setter keeps React's tracker "stale" until the event fires.
  const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  if (!nativeSetter) throw new Error("Missing HTMLInputElement.value setter");
  nativeSetter.call(input, value);
  input.dispatchEvent(new Event("input", { bubbles: true }));
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
      "Unpivot Columns",
      "Fill Down",
      "Replace Values",
      "Add Column",
    ]) {
      expect(host?.textContent).toContain(label);
    }
  });

  it("constructs a minimally-valid operation for menu selections", async () => {
    const preview = new DataTable([{ name: "  Region  ", type: "string" }], []);
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

  it("generates a unique Add Column name using existing query steps (even if preview schema is stale)", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query: Query = {
      ...baseQuery(),
      steps: [
        { id: "s1", name: "Added Custom", operation: { type: "addColumn", name: "  Custom  ", formula: "0" } },
      ],
    };
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query, preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Add Column").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({ type: "addColumn", name: "Custom 1", formula: "[Region]" });
  });

  it("constructs a minimally-valid unpivot operation", async () => {
    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Sales", type: "number" },
      ],
      [],
    );
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Unpivot Columns").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({
      type: "unpivot",
      columns: ["Sales"],
      nameColumn: "Attribute",
      valueColumn: "Value",
    });
  });

  it("generates unique unpivot output column names when the defaults already exist", async () => {
    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Sales", type: "number" },
      ],
      [],
    );
    const query: Query = {
      ...baseQuery(),
      steps: [
        {
          id: "s1",
          name: "Unpivoted",
          operation: { type: "unpivot", columns: ["Sales"], nameColumn: "Attribute", valueColumn: "Value" },
        },
      ],
    };
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query, preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Unpivot Columns").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({
      type: "unpivot",
      columns: ["Sales"],
      nameColumn: "Attribute 1",
      valueColumn: "Value 1",
    });
  });

  it("generates a unique Rename Column newName when a conflicting name already exists", async () => {
    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Region (Renamed)", type: "string" },
      ],
      [],
    );
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Rename Columns").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({
      type: "renameColumn",
      oldName: "Region",
      newName: "Region (Renamed) 1",
    });
  });

  it("considers existing renameColumn steps when generating a default Rename Column newName", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query: Query = {
      ...baseQuery(),
      steps: [
        { id: "s1", name: "Renamed", operation: { type: "renameColumn", oldName: "Region", newName: "Region (Renamed)" } },
      ],
    };
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query, preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Rename Columns").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({
      type: "renameColumn",
      oldName: "Region",
      newName: "Region (Renamed) 1",
    });
  });

  it("generates a unique Group By aggregation name when it would collide with a group column name", async () => {
    const preview = new DataTable(
      [
        { name: "Count Rows", type: "number" },
        { name: "Other", type: "string" },
      ],
      [],
    );
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Group By").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({
      type: "groupBy",
      groupColumns: ["Count Rows"],
      aggregations: [{ column: "Other", op: "count", as: "Count Rows 1" }],
    });
  });

  it("considers existing groupBy aggregation output names when generating a default Add Column name", async () => {
    const preview = new DataTable(
      [
        { name: "Region", type: "string" },
        { name: "Sales", type: "number" },
      ],
      [],
    );
    const query: Query = {
      ...baseQuery(),
      steps: [
        {
          id: "s1",
          name: "Grouped",
          operation: { type: "groupBy", groupColumns: ["Region"], aggregations: [{ column: "Sales", op: "count", as: "Custom" }] },
        },
      ],
    };
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query, preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      findButtonByText(host!, "Add Column").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({ type: "addColumn", name: "Custom 1", formula: "[Region]" });
  });

  it("disables schema-dependent operations when preview schema is missing", async () => {
    const onAddStep = vi.fn();
    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview: null }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const filterButton = findButtonByText(host!, "Filter Rows");
    expect(filterButton.disabled).toBe(true);
    expect(filterButton.title).toContain("Preview schema required");
    expect(host?.textContent).toContain("Preview schema required");

    const keepTopRowsButton = findButtonByText(host!, "Keep Top Rows");
    expect(keepTopRowsButton.disabled).toBe(false);

    await act(async () => {
      keepTopRowsButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    expect(onAddStep).toHaveBeenCalledWith({ type: "take", count: 100 });
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

  it("supports ArrowDown/ArrowUp navigation within the operation menu", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(5);
    });

    expect((document.activeElement as HTMLElement | null)?.textContent?.trim()).toBe("Keep Top Rows");

    await act(async () => {
      (document.activeElement as HTMLElement).dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    });

    expect((document.activeElement as HTMLElement | null)?.textContent?.trim()).toBe("Filter Rows");

    await act(async () => {
      (document.activeElement as HTMLElement).dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowUp", bubbles: true }));
    });

    expect((document.activeElement as HTMLElement | null)?.textContent?.trim()).toBe("Keep Top Rows");
  });

  it("allows pressing Enter on a focused menu item to apply it", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const onAddStep = vi.fn();

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={onAddStep} aiContext={{ query: baseQuery(), preview }} />);
    });

    await act(async () => {
      findButtonByText(host!, "+ Add step").dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(5);
    });

    expect((document.activeElement as HTMLElement | null)?.textContent?.trim()).toBe("Keep Top Rows");

    await act(async () => {
      (document.activeElement as HTMLElement).dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    });

    expect(onAddStep).toHaveBeenCalledWith({ type: "take", count: 100 });
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
      setTextInputValue(input, "   ");
    });
    expect(suggestButton.disabled).toBe(true);
    expect(onAiSuggest).not.toHaveBeenCalled();

    await act(async () => {
      setTextInputValue(input, "filter to non-null");
    });

    expect(suggestButton.disabled).toBe(false);

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAiSuggest).toHaveBeenCalledWith("filter to non-null", aiContext);
    expect(suggestButton.textContent).toContain("Suggesting");

    await act(async () => {
      deferred.resolve([
        { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } },
        { type: "addColumn", name: "Flag", formula: "1" },
      ]);
      await flushMicrotasks(10);
    });

    const suggestionButtons = host!.querySelectorAll("button.query-editor-add-step__suggestion");
    expect(suggestionButtons.length).toBe(2);
    expect(suggestionButtons[0]?.textContent).toBe("Filter Rows (Region)");
    expect(suggestionButtons[1]?.textContent).toBe("Add Column (Flag)");
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
      setTextInputValue(input, "filter");
    });
    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(10);
    });

    expect(host!.querySelectorAll("button.query-editor-add-step__suggestion").length).toBe(1);

    await act(async () => {
      setTextInputValue(input, "filter again");
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

  it("triggers AI suggest when pressing Enter in the intent input", async () => {
    const preview = new DataTable([{ name: "Region", type: "string" }], []);
    const query = baseQuery();
    const aiContext = { query, preview };

    const onAiSuggest = vi.fn(async () => []);

    await act(async () => {
      root?.render(<AddStepMenu onAddStep={() => {}} onAiSuggest={onAiSuggest} aiContext={aiContext} />);
    });

    const input = host!.querySelector("input.query-editor-add-step__ai-input") as HTMLInputElement;

    await act(async () => {
      setTextInputValue(input, "   hello  ");
    });

    await act(async () => {
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      await flushMicrotasks(10);
    });

    expect(onAiSuggest).toHaveBeenCalledWith("hello", aiContext);
    expect(host?.textContent).toContain("No suggestions.");
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
      setTextInputValue(input, "do something");
    });

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(5);
    });

    expect(host?.textContent).toContain("AI exploded");
    // Error state should not be confused with the empty suggestions state.
    expect(host?.textContent).not.toContain("No suggestions.");
  });

  it("ignores stale AI suggestions when the intent changes during a request", async () => {
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
      setTextInputValue(input, "filter to non-null");
    });

    await act(async () => {
      suggestButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onAiSuggest).toHaveBeenCalledTimes(1);
    expect(suggestButton.textContent).toContain("Suggesting");

    // Change intent while the request is in flight.
    await act(async () => {
      setTextInputValue(input, "something else");
      await flushMicrotasks(5);
    });

    expect(suggestButton.textContent).toContain("Suggest next");

    // Resolve the original request; it should be ignored.
    await act(async () => {
      deferred.resolve([{ type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } }]);
      await flushMicrotasks(10);
    });

    expect(host!.querySelectorAll("button.query-editor-add-step__suggestion").length).toBe(0);
  });
});
