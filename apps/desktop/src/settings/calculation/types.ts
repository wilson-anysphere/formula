export type CalculationMode = "automatic" | "automaticNoTable" | "manual";

export interface IterativeCalculationSettings {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
}

export interface CalcSettings {
  calculationMode: CalculationMode;
  calculateBeforeSave: boolean;
  /**
   * When `true`, calculations use full floating point precision.
   *
   * When `false`, the workbook is in "precision as displayed" mode (Excel's
   * "Set precision as displayed").
   */
  fullPrecision: boolean;
  iterative: IterativeCalculationSettings;
}
