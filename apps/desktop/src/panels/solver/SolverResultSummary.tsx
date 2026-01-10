import React from "react";

import { t } from "../../i18n/index.js";

import type { SolverOutcome } from "./types";

type Props = {
  outcome: SolverOutcome;
  onKeep: () => void;
  onRestore: () => void;
};

export function SolverResultSummary({ outcome, onKeep, onRestore }: Props) {
  return (
    <div className="solver-results">
      <h3>{t("solver.results.title")}</h3>

      <dl>
        <dt>{t("solver.results.status")}</dt>
        <dd>{outcome.status}</dd>
        <dt>{t("solver.results.iterations")}</dt>
        <dd>{outcome.iterations}</dd>
        <dt>{t("solver.results.objective")}</dt>
        <dd>{outcome.bestObjective}</dd>
        <dt>{t("solver.results.maxConstraintViolation")}</dt>
        <dd>{outcome.maxConstraintViolation}</dd>
      </dl>

      <div style={{ display: "flex", gap: 12 }}>
        <button type="button" onClick={onKeep}>
          {t("solver.results.keepSolution")}
        </button>
        <button type="button" onClick={onRestore}>
          {t("solver.results.restoreOriginalValues")}
        </button>
      </div>
    </div>
  );
}
