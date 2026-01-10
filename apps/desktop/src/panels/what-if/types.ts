export type CellRef = string;

// Goal Seek
export interface GoalSeekParams {
  targetCell: CellRef;
  targetValue: number;
  changingCell: CellRef;
  maxIterations?: number;
  tolerance?: number;
}

export type GoalSeekStatus =
  | "Converged"
  | "MaxIterationsReached"
  | "NoBracketFound"
  | "NumericalFailure";

export interface GoalSeekResult {
  status: GoalSeekStatus;
  solution: number;
  iterations: number;
  finalOutput: number;
  finalError: number;
}

export interface GoalSeekProgress {
  iteration: number;
  input: number;
  output: number;
  error: number;
}

export type WhatIfCellValue =
  | { type: "number"; value: number }
  | { type: "text"; value: string }
  | { type: "bool"; value: boolean }
  | { type: "blank" };

// Scenario Manager
export interface Scenario {
  id: number;
  name: string;
  changingCells: CellRef[];
  values: Record<CellRef, WhatIfCellValue>;
  createdMs?: number;
  createdBy: string;
  comment?: string;
}

export interface SummaryReport {
  changingCells: CellRef[];
  resultCells: CellRef[];
  results: Record<string, Record<CellRef, WhatIfCellValue>>;
}

// Monte Carlo
export type Distribution =
  | { type: "normal"; mean: number; stdDev: number }
  | { type: "uniform"; min: number; max: number }
  | { type: "triangular"; min: number; mode: number; max: number }
  | { type: "lognormal"; mean: number; stdDev: number }
  | { type: "discrete"; values: number[]; probabilities: number[] }
  | { type: "beta"; alpha: number; beta: number; min?: number; max?: number }
  | { type: "exponential"; rate: number }
  | { type: "poisson"; lambda: number };

export interface InputDistribution {
  cell: CellRef;
  distribution: Distribution;
}

export interface CorrelationMatrix {
  matrix: number[][];
}

export interface SimulationConfig {
  iterations: number;
  inputDistributions: InputDistribution[];
  outputCells: CellRef[];
  seed?: number;
  correlations?: CorrelationMatrix;
  histogramBins?: number;
}

export interface HistogramBin {
  start: number;
  end: number;
  count: number;
}

export interface OutputStatistics {
  mean: number;
  median: number;
  stdDev: number;
  min: number;
  max: number;
  percentiles: Record<string, number>;
  histogram: { bins: HistogramBin[] };
}

export interface SimulationResult {
  iterations: number;
  outputStats: Record<CellRef, OutputStatistics>;
  outputSamples: Record<CellRef, number[]>;
}

export interface SimulationProgress {
  completedIterations: number;
  totalIterations: number;
}

/**
 * Bridge used by UI components to talk to the backend (Tauri/IPC).
 *
 * The desktop app will provide an implementation that calls into the Rust
 * engine and streams progress updates.
 */
export interface WhatIfApi {
  goalSeek: (params: GoalSeekParams, onProgress?: (p: GoalSeekProgress) => void) => Promise<GoalSeekResult>;
  listScenarios: () => Promise<Scenario[]>;
  createScenario: (scenario: Omit<Scenario, "id">) => Promise<Scenario>;
  applyScenario: (scenarioId: number) => Promise<void>;
  restoreBaseScenario: () => Promise<void>;
  generateSummaryReport: (resultCells: CellRef[], scenarioIds: number[]) => Promise<SummaryReport>;
  runMonteCarlo: (
    config: SimulationConfig,
    onProgress?: (p: SimulationProgress) => void
  ) => Promise<SimulationResult>;
}
