import React, { useEffect, useMemo, useRef, useState } from "react";

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
  const [invalidField, setInvalidField] = useState<"targetValue" | null>(null);

  const parsedTargetValue = useMemo(() => Number(targetValue), [targetValue]);
  const reactInstanceId = React.useId();
  const domInstanceId = useMemo(() => reactInstanceId.replace(/[^a-zA-Z0-9_-]/g, "-"), [reactInstanceId]);
  const titleId = useMemo(() => `goal-seek-title-${domInstanceId}`, [domInstanceId]);
  const errorId = useMemo(() => `goal-seek-error-${domInstanceId}`, [domInstanceId]);
  const targetCellRef = useRef<HTMLInputElement | null>(null);
  const dialogRef = useRef<HTMLDivElement | null>(null);

  const trapTab = (event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== "Tab") return;
    const root = dialogRef.current;
    if (!root) return;

    const focusables = Array.from(
      root.querySelectorAll<HTMLElement>("button, [href], input, select, textarea, [tabindex]"),
    ).filter((el) => {
      if (el.getAttribute("aria-hidden") === "true") return false;
      // Ignore non-tabbable elements and explicitly disabled controls.
      if (el.getAttribute("tabindex") === "-1") return false;
      if ((el as HTMLButtonElement).disabled) return false;
      return true;
    });

    if (focusables.length === 0) return;
    const first = focusables[0]!;
    const last = focusables[focusables.length - 1]!;
    const active = document.activeElement as HTMLElement | null;

    if (event.shiftKey) {
      if (active === first) {
        event.preventDefault();
        last.focus();
      }
      return;
    }

    if (active === last) {
      event.preventDefault();
      first.focus();
    }
  };

  useEffect(() => {
    if (!open) return;
    // Focus the first input so keyboard users can immediately type.
    targetCellRef.current?.focus();
  }, [open]);

  async function run() {
    setError(null);
    setInvalidField(null);
    setResult(null);
    setProgress(null);

    if (!Number.isFinite(parsedTargetValue)) {
      setError(t("whatIf.goalSeek.error.targetMustBeNumber"));
      setInvalidField("targetValue");
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
      aria-labelledby={titleId}
      aria-describedby={error ? errorId : undefined}
      aria-busy={running ? true : undefined}
      data-keybinding-barrier="true"
      data-testid="goal-seek-dialog"
      ref={dialogRef}
      onKeyDown={(event) => {
        if (event.key !== "Escape") return;
        if (running) return;
        event.preventDefault();
        event.stopPropagation();
        onClose();
      }}
      onKeyDownCapture={(event) => {
        // Keep tab focus inside the dialog while it is open.
        trapTab(event);
      }}
    >
      <div className="what-if-dialog__header">
        <h3 className="dialog__title" id={titleId}>
          {t("whatIf.goalSeek.title")}
        </h3>
        <button type="button" className="what-if__button" onClick={onClose} disabled={running}>
          {t("whatIf.goalSeek.close")}
        </button>
      </div>

      <div className="what-if-dialog__content">
        <div className="what-if-grid">
          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.setCell")}</span>
            <input
              className="what-if__input what-if__input--mono"
              value={targetCell}
              onChange={(e) => setTargetCell(e.target.value)}
              disabled={running}
              spellCheck={false}
              autoCapitalize="off"
              ref={targetCellRef}
            />
          </label>

          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.toValue")}</span>
            <input
              className="what-if__input"
              value={targetValue}
              onChange={(e) => {
                setTargetValue(e.target.value);
                if (invalidField === "targetValue") {
                  setInvalidField(null);
                  setError(null);
                }
              }}
              disabled={running}
              inputMode="decimal"
              aria-invalid={invalidField === "targetValue" ? true : undefined}
              aria-describedby={invalidField === "targetValue" ? errorId : undefined}
            />
          </label>

          <label className="what-if__field">
            <span className="what-if__label">{t("whatIf.goalSeek.byChangingCell")}</span>
            <input
              className="what-if__input what-if__input--mono"
              value={changingCell}
              onChange={(e) => setChangingCell(e.target.value)}
              disabled={running}
              spellCheck={false}
              autoCapitalize="off"
            />
          </label>

          <div className="what-if__actions">
            <button type="button" className="what-if__button what-if__button--primary" onClick={run} disabled={running}>
              {running ? t("whatIf.goalSeek.running") : t("whatIf.goalSeek.solve")}
            </button>
          </div>
        </div>

        {error ? (
          <p className="what-if__message what-if__message--error" role="alert" id={errorId}>
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
