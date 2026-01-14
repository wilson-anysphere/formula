// @vitest-environment jsdom

import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { GoalSeekDialog } from "./GoalSeekDialog";
import { MonteCarloWizard } from "./MonteCarloWizard";
import { ScenarioManagerPanel } from "./ScenarioManagerPanel";
import type { GoalSeekParams, GoalSeekResult, Scenario, SimulationConfig, SimulationResult, SummaryReport, WhatIfApi } from "./types";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

function normalizeHtmlForSnapshot(html: string): string {
  // React.useId() produces globally incrementing ids, so snapshotting raw HTML can
  // churn if other tests add components that call useId(). Normalize the known
  // ids we generate in these components so snapshots stay stable.
  return html
    .replace(/goal-seek-title-[^"\\s]+/g, "goal-seek-title-<id>")
    .replace(/goal-seek-error-[^"\\s]+/g, "goal-seek-error-<id>")
    .replace(/scenario-manager-error-[^"\\s]+/g, "scenario-manager-error-<id>")
    .replace(/monte-carlo-error-[^"\\s]+/g, "monte-carlo-error-<id>");
}

function setTextInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
  if (setter) {
    setter.call(input, value);
  } else {
    input.value = value;
  }

  // React attaches internal value trackers to inputs; dispatching events after
  // using the native setter ensures onChange sees the update in jsdom.
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function createStubApi(overrides: Partial<WhatIfApi> = {}): WhatIfApi {
  const notImplemented = async () => {
    throw new Error("Not implemented in test stub");
  };

  return {
    goalSeek: overrides.goalSeek ?? (async (_params: GoalSeekParams) => ({ status: "Converged", solution: 1, iterations: 1, finalOutput: 1, finalError: 0 } satisfies GoalSeekResult)),
    listScenarios: overrides.listScenarios ?? (async () => []),
    createScenario: overrides.createScenario ?? notImplemented,
    applyScenario: overrides.applyScenario ?? (async () => {}),
    restoreBaseScenario: overrides.restoreBaseScenario ?? (async () => {}),
    generateSummaryReport:
      overrides.generateSummaryReport ??
      (async () =>
        ({
          changingCells: [],
          resultCells: [],
          results: {},
        }) satisfies SummaryReport),
    runMonteCarlo:
      overrides.runMonteCarlo ??
      (async (_config: SimulationConfig) =>
        ({
          iterations: 0,
          outputStats: {},
          outputSamples: {},
        }) satisfies SimulationResult),
  };
}

describe("what-if UI components", () => {
  afterEach(() => {
    document.body.innerHTML = "";
  });

  it("ScenarioManagerPanel renders with token-driven class structure (no inline styles)", async () => {
    const scenarios: Scenario[] = [
      {
        id: 1,
        name: "Base",
        changingCells: ["A1"],
        values: { A1: { type: "number", value: 1 } },
        createdBy: "tester",
      },
      {
        id: 2,
        name: "Optimistic",
        changingCells: ["A1", "A2"],
        values: { A1: { type: "number", value: 2 }, A2: { type: "text", value: "ok" } },
        createdBy: "tester",
        comment: "demo",
      },
    ];

    const api = createStubApi({
      listScenarios: vi.fn(async () => scenarios),
    });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(ScenarioManagerPanel, { api }));
      await flushPromises();
    });

    expect(host.querySelectorAll("[style]").length).toBe(0);
    expect(host.querySelector('[data-testid="scenario-manager-panel"]')).toBeTruthy();
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("ScenarioManagerPanel renders selected scenario metadata and a summary report table", async () => {
    const scenarios: Scenario[] = [
      {
        id: 1,
        name: "Base",
        changingCells: ["A1"],
        values: { A1: { type: "number", value: 1 } },
        createdBy: "tester",
      },
      {
        id: 2,
        name: "Optimistic",
        changingCells: ["A1", "A2"],
        values: { A1: { type: "number", value: 2 }, A2: { type: "text", value: "ok" } },
        createdBy: "tester",
        comment: "demo",
      },
    ];

    const generateSummaryReport = vi.fn(async () => {
      return {
        changingCells: ["A1"],
        resultCells: ["B1", "C1"],
        results: {
          Base: {
            B1: { type: "number", value: 10 },
            C1: { type: "bool", value: true },
          },
          Optimistic: {
            B1: { type: "text", value: "hi" },
            C1: { type: "blank" },
          },
        },
      } satisfies SummaryReport;
    });

    const api = createStubApi({
      listScenarios: vi.fn(async () => scenarios),
      generateSummaryReport,
    });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(ScenarioManagerPanel, { api }));
      await flushPromises();
    });

    const select = host.querySelector("select.what-if__select") as HTMLSelectElement | null;
    expect(select).toBeTruthy();
    await act(async () => {
      select!.value = "2";
      select!.dispatchEvent(new Event("change", { bubbles: true }));
    });

    const resultCellsInput = host.querySelector("input.what-if__input") as HTMLInputElement | null;
    expect(resultCellsInput).toBeTruthy();
    await act(async () => {
      setTextInputValue(resultCellsInput!, "B1, C1");
    });

    const summaryBtn = Array.from(host.querySelectorAll("button")).find((btn) => btn.textContent === "Summary Report") as
      | HTMLButtonElement
      | undefined;
    expect(summaryBtn).toBeTruthy();

    await act(async () => {
      summaryBtn!.click();
      await flushPromises();
    });

    expect(generateSummaryReport).toHaveBeenCalledWith(["B1", "C1"], [1, 2]);
    expect(host.querySelector('[data-testid="scenario-manager-selected-meta"]')?.textContent).toContain("demo");
    expect(host.querySelector('[data-testid="scenario-manager-report"]')).toBeTruthy();
    expect(host.querySelectorAll("[style]").length).toBe(0);
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("GoalSeekDialog renders a dialog shell with labeled fields (no inline styles)", async () => {
    const api = createStubApi();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(GoalSeekDialog, { api, open: true, onClose: () => {} }));
    });

    expect(host.querySelectorAll("[style]").length).toBe(0);
    const dialog = host.querySelector('[data-testid="goal-seek-dialog"]');
    expect(dialog).toBeTruthy();
    expect(dialog?.getAttribute("role")).toBe("dialog");
    expect(dialog?.getAttribute("aria-modal")).toBe("true");
    expect(dialog?.querySelectorAll("label").length).toBeGreaterThanOrEqual(3);
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("GoalSeekDialog shows progress + result after running", async () => {
    const goalSeek = vi.fn(async (_params: GoalSeekParams, onProgress?: (p: any) => void) => {
      onProgress?.({ iteration: 1, input: 0.5, output: 0.2, error: -0.8 });
      await flushPromises();
      onProgress?.({ iteration: 2, input: 1, output: 1, error: 0 });
      return { status: "Converged", solution: 1, iterations: 2, finalOutput: 1, finalError: 0 } satisfies GoalSeekResult;
    });

    const api = createStubApi({ goalSeek });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(GoalSeekDialog, { api, open: true, onClose: () => {} }));
    });

    const solveBtn = host.querySelector("button.what-if__button--primary") as HTMLButtonElement | null;
    expect(solveBtn).toBeTruthy();

    await act(async () => {
      solveBtn!.click();
      await flushPromises();
    });

    expect(goalSeek).toHaveBeenCalled();
    expect(host.querySelector('[data-testid="goal-seek-progress"]')).toBeTruthy();
    expect(host.querySelector('[data-testid="goal-seek-result"]')).toBeTruthy();
    expect(host.querySelectorAll("[style]").length).toBe(0);
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("MonteCarloWizard renders responsive rows with accessible input labels (no inline styles)", async () => {
    const api = createStubApi();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(MonteCarloWizard, { api }));
    });

    expect(host.querySelectorAll("[style]").length).toBe(0);
    expect(host.querySelector('[data-testid="monte-carlo-wizard"]')).toBeTruthy();
    expect(host.querySelector('input[aria-label="Input cell"]')).toBeTruthy();
    expect(host.querySelector('select[aria-label="Distribution type"]')).toBeTruthy();
    expect(host.querySelector('input[aria-label="Distribution JSON"]')).toBeTruthy();
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("MonteCarloWizard shows progress + results after running", async () => {
    const runMonteCarlo = vi.fn(async (_config: SimulationConfig, onProgress?: (p: any) => void) => {
      onProgress?.({ completedIterations: 10, totalIterations: 1000 });
      await flushPromises();
      onProgress?.({ completedIterations: 1000, totalIterations: 1000 });
      return {
        iterations: 1000,
        outputStats: {
          B1: {
            mean: 1,
            median: 1,
            stdDev: 0.5,
            min: 0,
            max: 2,
            percentiles: { "5": 0.2, "95": 1.8 },
            histogram: { bins: [] },
          },
        },
        outputSamples: { B1: [] },
      } satisfies SimulationResult;
    });

    const api = createStubApi({ runMonteCarlo });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(MonteCarloWizard, { api }));
    });

    const runBtn = Array.from(host.querySelectorAll("button")).find((btn) => btn.textContent === "Run simulation") as
      | HTMLButtonElement
      | undefined;
    expect(runBtn).toBeTruthy();

    await act(async () => {
      runBtn!.click();
      await flushPromises();
    });

    expect(runMonteCarlo).toHaveBeenCalled();
    expect(host.querySelector('[data-testid="monte-carlo-progress"]')).toBeTruthy();
    expect(host.querySelector('[data-testid="monte-carlo-results"]')).toBeTruthy();
    expect(host.querySelectorAll("[style]").length).toBe(0);
    expect(normalizeHtmlForSnapshot(host.innerHTML)).toMatchSnapshot();
  });

  it("MonteCarloWizard marks invalid distribution JSON only after run and clears when fixed", async () => {
    const runMonteCarlo = vi.fn(async () => {
      throw new Error("Should not run when distribution JSON is invalid");
    });

    const api = createStubApi({ runMonteCarlo });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(React.createElement(MonteCarloWizard, { api }));
    });

    const jsonInput = host.querySelector('input[aria-label="Distribution JSON"]') as HTMLInputElement | null;
    expect(jsonInput).toBeTruthy();

    await act(async () => {
      setTextInputValue(jsonInput!, "{");
    });

    // Still editable (no red invalid state) until the user tries to run.
    expect(jsonInput?.getAttribute("aria-invalid")).toBe(null);

    const runBtn = Array.from(host.querySelectorAll("button")).find((btn) => btn.textContent === "Run simulation") as
      | HTMLButtonElement
      | undefined;
    expect(runBtn).toBeTruthy();

    await act(async () => {
      runBtn!.click();
      await flushPromises();
    });

    expect(runMonteCarlo).not.toHaveBeenCalled();

    const alert = host.querySelector('[role="alert"]') as HTMLElement | null;
    expect(alert).toBeTruthy();
    expect(alert?.textContent).toContain("Fix invalid distribution JSON before running.");
    expect(jsonInput?.getAttribute("aria-invalid")).toBe("true");
    expect(jsonInput?.getAttribute("aria-describedby")).toBe(alert?.id ?? null);

    await act(async () => {
      setTextInputValue(jsonInput!, '{"type":"normal","mean":0,"stdDev":1}');
      await flushPromises();
    });

    expect(host.querySelector('[role="alert"]')).toBeNull();
    expect(jsonInput?.getAttribute("aria-invalid")).toBe(null);
    expect(jsonInput?.getAttribute("aria-describedby")).toBe(null);
  });
});
