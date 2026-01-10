import React from "react";

import { t } from "../../i18n/index.js";

import type { SolverProgress } from "./types";

type Props = {
  progress: SolverProgress | null;
  onCancel: () => void;
};

export function SolverProgressView({ progress, onCancel }: Props) {
  return (
    <div className="solver-progress">
      <h3>{t("solver.progress.title")}</h3>
      {progress ? (
        <dl>
          <dt>{t("solver.progress.iteration")}</dt>
          <dd>{progress.iteration}</dd>
          <dt>{t("solver.progress.bestObjective")}</dt>
          <dd>{progress.bestObjective}</dd>
          <dt>{t("solver.progress.currentObjective")}</dt>
          <dd>{progress.currentObjective}</dd>
          <dt>{t("solver.progress.maxConstraintViolation")}</dt>
          <dd>{progress.maxConstraintViolation}</dd>
        </dl>
      ) : (
        <p>{t("solver.progress.starting")}</p>
      )}

      <button type="button" onClick={onCancel}>
        {t("solver.progress.cancel")}
      </button>
    </div>
  );
}
