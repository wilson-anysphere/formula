import React, { useMemo, useState } from "react";

import type { GoalSeekParams, GoalSeekProgress, GoalSeekResult, WhatIfApi } from "./types";
import { t } from "../../i18n/index.js";

export interface GoalSeekDialogProps {
  api: WhatIfApi;
  open: boolean;
  onClose: () => void;
}

export function GoalSeekDialog({ api, open, onClose }: GoalSeekDialogProps) {
  const [targetCell, setTargetCell] = useState("B1");
  const [changingCell, setChangingCell] = useState("A1");
  const [targetValue, setTargetValue] = useState("0");

  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<GoalSeekProgress | null>(null);
  const [result, setResult] = useState<GoalSeekResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const parsedTargetValue = useMemo(() => Number(targetValue), [targetValue]);

  async function run() {
    setError(null);
    setResult(null);
    setProgress(null);

    if (!Number.isFinite(parsedTargetValue)) {
      setError(t("whatIf.goalSeek.error.targetMustBeNumber"));
      return;
    }

    const params: GoalSeekParams = {
      targetCell: targetCell.trim(),
      changingCell: changingCell.trim(),
      targetValue: parsedTargetValue,
    };

    setRunning(true);
    try {
      const res = await api.goalSeek(params, (p) => setProgress(p));
      setResult(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setRunning(false);
    }
  }

  if (!open) return null;

  return (
    <div
      className="dialog what-if-dialog"
      role="dialog"
      aria-modal="true"
      aria-label={t("whatIf.goalSeek.title")}
      data-testid="goal-seek-dialog"
    >
      <div className="what-if-dialog__header">
        <h3 className="dialog__title">{t("whatIf.goalSeek.title")}</h3>
        <button type="button" className="what-if__button" onClick={onClose} disabled={running}>
          {t("whatIf.goalSeek.close")}
        </button>
      </div>

      <div className="what-if-dialog__content">
        <div className="what-if-grid">
          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.setCell")}</span>
            <input className="what-if__input" value={targetCell} onChange={(e) => setTargetCell(e.target.value)} disabled={running} />
          </label>

          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.toValue")}</span>
            <input className="what-if__input" value={targetValue} onChange={(e) => setTargetValue(e.target.value)} disabled={running} />
          </label>

          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.byChangingCell")}</span>
            <input className="what-if__input" value={changingCell} onChange={(e) => setChangingCell(e.target.value)} disabled={running} />
          </label>

          <div className="what-if__actions">
            <button type="button" className="what-if__button what-if__button--primary" onClick={run} disabled={running}>
              {running ? t("whatIf.goalSeek.running") : t("whatIf.goalSeek.solve")}
            </button>
          </div>
        </div>

        {error ? (
          <p className="what-if__message what-if__message--error" role="alert">
            {error}
          </p>
        ) : null}

        {progress ? (
          <div className="what-if__mono-block" role="status" data-testid="goal-seek-progress">
            <div>
              {t("whatIf.goalSeek.progress.iteration")}: {progress.iteration}
            </div>
            <div>
              {t("whatIf.goalSeek.progress.input")}: {progress.input}
            </div>
            <div>
              {t("whatIf.goalSeek.progress.output")}: {progress.output}
            </div>
            <div>
              {t("whatIf.goalSeek.progress.error")}: {progress.error}
            </div>
          </div>
        ) : null}

        {result ? (
          <div data-testid="goal-seek-result">
            <h4 className="what-if__section-title">{t("whatIf.goalSeek.result.title")}</h4>
            <div className="what-if__mono-block">
              <div>
                {t("whatIf.goalSeek.result.status")}: {t(`whatIf.goalSeek.status.${result.status}`)}
              </div>
              <div>
                {t("whatIf.goalSeek.result.solution")}: {result.solution}
              </div>
              <div>
                {t("whatIf.goalSeek.result.iterations")}: {result.iterations}
              </div>
              <div>
                {t("whatIf.goalSeek.result.finalOutput")}: {result.finalOutput}
              </div>
              <div>
                {t("whatIf.goalSeek.result.finalError")}: {result.finalError}
              </div>
            </div>
          </div>
        ) : null}
      </div>
    </div>
  );
}
