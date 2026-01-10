export type SolveMethod = "simplex" | "grg" | "evolutionary";
export type ObjectiveKind = "maximize" | "minimize" | "target";

export type Relation = "<=" | ">=" | "=";
export type VarType = "continuous" | "integer" | "binary";

export interface SolverVariableSpec {
  /** Spreadsheet address / named range identifying the decision variable cell. */
  ref: string;
  lower?: number;
  upper?: number;
  type: VarType;
}

export interface SolverConstraintSpec {
  /** Spreadsheet address / named range identifying the LHS value cell. */
  ref: string;
  relation: Relation;
  rhs: number;
  tolerance?: number;
}

export interface SolverConfig {
  method: SolveMethod;

  objectiveRef: string;
  objectiveKind: ObjectiveKind;
  targetValue?: number;
  targetTolerance?: number;

  variables: SolverVariableSpec[];
  constraints: SolverConstraintSpec[];

  maxIterations?: number;
  tolerance?: number;
}

export interface SolverProgress {
  iteration: number;
  bestObjective: number;
  currentObjective: number;
  maxConstraintViolation: number;
}

export type SolveStatus =
  | "optimal"
  | "feasible"
  | "infeasible"
  | "unbounded"
  | "iterationLimit"
  | "cancelled";

export interface SolverOutcome {
  status: SolveStatus;
  iterations: number;
  originalVars: number[];
  bestVars: number[];
  bestObjective: number;
  maxConstraintViolation: number;
}

