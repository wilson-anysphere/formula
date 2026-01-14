// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";

import type { QueryStep } from "@formula/power-query";

import { setLocale } from "../../../i18n/index.js";
import { StepsList } from "./StepsList";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

describe("StepsList", () => {
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

  it("falls back to operation label when step.name is missing or equals op.type", async () => {
    const steps: QueryStep[] = [
      {
        id: "s1",
        name: "  filterRows  ",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "isNotNull" } },
      },
      {
        id: "s2",
        name: "",
        operation: { type: "sortRows", sortBy: [{ column: "Region", direction: "ascending" }] },
      },
      {
        id: "s3",
        name: "  My Custom Step  ",
        operation: { type: "removeColumns", columns: ["Region"] },
      },
    ];

    await act(async () => {
      root?.render(<StepsList steps={steps} selectedIndex={0} onSelect={() => {}} />);
    });

    const buttons = Array.from(host?.querySelectorAll<HTMLButtonElement>(".query-editor-steps__button") ?? []);
    expect(buttons.map((b) => b.textContent)).toEqual(["Filter Rows (Region)", "Sort (Region)", "My Custom Step"]);
  });
});
