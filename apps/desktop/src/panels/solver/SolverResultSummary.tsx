import React from "react";

import type { SolverOutcome } from "./types";

type Props = {
  outcome: SolverOutcome;
  onKeep: () => void;
  onRestore: () => void;
};

export function SolverResultSummary({ outcome, onKeep, onRestore }: Props) {
  return (
    <div className="solver-results">
      <h3>Solver Results</h3>

      <dl>
        <dt>Status</dt>
        <dd>{outcome.status}</dd>
        <dt>Iterations</dt>
        <dd>{outcome.iterations}</dd>
        <dt>Objective</dt>
        <dd>{outcome.bestObjective}</dd>
        <dt>Max constraint violation</dt>
        <dd>{outcome.maxConstraintViolation}</dd>
      </dl>

      <div style={{ display: "flex", gap: 12 }}>
        <button type="button" onClick={onKeep}>
          Keep Solution
        </button>
        <button type="button" onClick={onRestore}>
          Restore Original Values
        </button>
      </div>
    </div>
  );
}

