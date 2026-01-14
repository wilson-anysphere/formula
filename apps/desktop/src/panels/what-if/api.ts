import type {
  GoalSeekParams,
  GoalSeekProgress,
  GoalSeekResult,
  Scenario,
  SimulationConfig,
  SimulationProgress,
  SimulationResult,
  SummaryReport,
  WhatIfApi,
} from "./types";

/**
 * Desktop What-If Analysis IPC stub.
 *
 * The real application should call into the Rust formula engine (Tauri command / MessageChannel).
 * We keep this API surface so UI panels/dialogs can be mounted without blocking on the IPC layer.
 */
export function createWhatIfApi(): WhatIfApi {
  const notWired = (feature: string) => new Error(`${feature} is not wired to the backend yet`);

  return {
    goalSeek: async (_params: GoalSeekParams, _onProgress?: (p: GoalSeekProgress) => void): Promise<GoalSeekResult> => {
      throw notWired("Goal Seek");
    },
    listScenarios: async (): Promise<Scenario[]> => [],
    createScenario: async (): Promise<Scenario> => {
      throw notWired("Scenario Manager");
    },
    applyScenario: async (): Promise<void> => {
      throw notWired("Scenario Manager");
    },
    restoreBaseScenario: async (): Promise<void> => {
      throw notWired("Scenario Manager");
    },
    generateSummaryReport: async (_resultCells: string[], _scenarioIds: number[]): Promise<SummaryReport> => {
      throw notWired("Scenario Manager");
    },
    runMonteCarlo: async (
      _config: SimulationConfig,
      _onProgress?: (p: SimulationProgress) => void,
    ): Promise<SimulationResult> => {
      throw notWired("Monte Carlo");
    },
  };
}
