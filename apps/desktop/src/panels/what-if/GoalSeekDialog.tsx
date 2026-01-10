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
    <div style={{ padding: 16, border: "1px solid var(--dialog-border)", borderRadius: 8, maxWidth: 420 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h3 style={{ margin: 0 }}>{t("whatIf.goalSeek.title")}</h3>
        <button onClick={onClose} disabled={running}>
          {t("whatIf.goalSeek.close")}
        </button>
      </div>

      <div style={{ display: "grid", gap: 8, marginTop: 12 }}>
        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.goalSeek.setCell")}</span>
          <input value={targetCell} onChange={(e) => setTargetCell(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.goalSeek.toValue")}</span>
          <input value={targetValue} onChange={(e) => setTargetValue(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.goalSeek.byChangingCell")}</span>
          <input value={changingCell} onChange={(e) => setChangingCell(e.target.value)} disabled={running} />
        </label>

        <button onClick={run} disabled={running}>
          {running ? t("whatIf.goalSeek.running") : t("whatIf.goalSeek.solve")}
        </button>
      </div>

      {error ? <p style={{ color: "var(--error)" }}>{error}</p> : null}

      {progress ? (
        <div style={{ marginTop: 12, fontFamily: "monospace", fontSize: 12 }}>
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
        <div style={{ marginTop: 12 }}>
          <h4 style={{ margin: "8px 0" }}>{t("whatIf.goalSeek.result.title")}</h4>
          <div style={{ fontFamily: "monospace", fontSize: 12 }}>
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
  );
}
