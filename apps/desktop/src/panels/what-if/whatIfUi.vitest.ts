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
    expect(host.innerHTML).toMatchSnapshot();
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
    expect(host.innerHTML).toMatchSnapshot();
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
    expect(host.innerHTML).toMatchSnapshot();
  });
});

