export const FORMULA_AUDITING_RIBBON_COMMAND_IDS = [
  "formulas.formulaAuditing.tracePrecedents",
  "formulas.formulaAuditing.traceDependents",
  "formulas.formulaAuditing.removeArrows",
] as const;

export type FormulaAuditingRibbonCommandId = (typeof FORMULA_AUDITING_RIBBON_COMMAND_IDS)[number];

