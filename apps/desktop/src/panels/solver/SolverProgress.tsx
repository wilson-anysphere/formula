import React from "react";

import type { SolverProgress } from "./types";

type Props = {
  progress: SolverProgress | null;
  onCancel: () => void;
};

export function SolverProgressView({ progress, onCancel }: Props) {
  return (
    <div className="solver-progress">
      <h3>Solving…</h3>
      {progress ? (
        <dl>
          <dt>Iteration</dt>
          <dd>{progress.iteration}</dd>
          <dt>Best objective</dt>
          <dd>{progress.bestObjective}</dd>
          <dt>Current objective</dt>
          <dd>{progress.currentObjective}</dd>
          <dt>Max constraint violation</dt>
          <dd>{progress.maxConstraintViolation}</dd>
        </dl>
      ) : (
        <p>Starting…</p>
      )}

      <button type="button" onClick={onCancel}>
        Cancel
      </button>
    </div>
  );
}

