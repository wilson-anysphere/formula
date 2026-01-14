// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DataTable, type Query, type QueryEngine, type QueryOperation, type QueryStep } from "@formula/power-query";

import { setLocale } from "../../i18n/index.js";
import { QueryEditorPanel } from "./QueryEditorPanel";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

async function flushMicrotasks(count = 8): Promise<void> {
  for (let i = 0; i < count; i++) await Promise.resolve();
}

function baseQuery(steps: QueryStep[]): Query {
  return { id: "q1", name: "Query 1", source: { type: "range", range: { values: [] } }, steps };
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

describe("QueryEditorPanel", () => {
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
  });

  it("passes AI suggest context for the currently-selected step (not the full query)", async () => {
    const steps: QueryStep[] = [
      { id: "s1", name: "Step 1", operation: { type: "take", count: 5 } },
      { id: "s2", name: "Step 2", operation: { type: "distinctRows", columns: null } },
    ];

    const engine = {
      executeQuery: vi.fn(async () => new DataTable([{ name: "Region", type: "string" }], [])),
    } as unknown as QueryEngine;

    const onAiSuggestNextSteps = vi.fn(async () => [] as QueryOperation[]);

    await act(async () => {
      root?.render(<QueryEditorPanel query={baseQuery(steps)} engine={engine} onAiSuggestNextSteps={onAiSuggestNextSteps} />);
      await flushMicrotasks(10);
    });

    // Select the first step in the list.
    await act(async () => {
      findButtonByText(host!, "Step 1").dispatchEvent(new MouseEvent("click", { bubbles: true }));
      await flushMicrotasks(10);
    });

    const input = host!.querySelector("input.query-editor-add-step__ai-input") as HTMLInputElement;
    expect(input).toBeTruthy();

    await act(async () => {
      input.value = "hello";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    });

    expect(onAiSuggestNextSteps).toHaveBeenCalledTimes(1);
    const [intent, ctx] = onAiSuggestNextSteps.mock.calls[0]!;
    expect(intent).toBe("hello");
    expect(ctx.query.steps).toHaveLength(1);
    expect(ctx.query.steps[0]?.id).toBe("s1");
  });
});
