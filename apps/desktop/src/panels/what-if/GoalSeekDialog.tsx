import React, { useMemo, useState } from "react";

import type { GoalSeekParams, GoalSeekProgress, GoalSeekResult, WhatIfApi } from "./types";

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
      setError("Target value must be a number.");
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
    <div style={{ padding: 16, border: "1px solid #ccc", borderRadius: 8, maxWidth: 420 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h3 style={{ margin: 0 }}>Goal Seek</h3>
        <button onClick={onClose} disabled={running}>
          Close
        </button>
      </div>

      <div style={{ display: "grid", gap: 8, marginTop: 12 }}>
        <label style={{ display: "grid", gap: 4 }}>
          <span>Set cell</span>
          <input value={targetCell} onChange={(e) => setTargetCell(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>To value</span>
          <input value={targetValue} onChange={(e) => setTargetValue(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>By changing cell</span>
          <input value={changingCell} onChange={(e) => setChangingCell(e.target.value)} disabled={running} />
        </label>

        <button onClick={run} disabled={running}>
          {running ? "Running…" : "Solve"}
        </button>
      </div>

      {error ? <p style={{ color: "crimson" }}>{error}</p> : null}

      {progress ? (
        <div style={{ marginTop: 12, fontFamily: "monospace", fontSize: 12 }}>
          <div>Iteration: {progress.iteration}</div>
          <div>Input: {progress.input}</div>
          <div>Output: {progress.output}</div>
          <div>Error: {progress.error}</div>
        </div>
      ) : null}

      {result ? (
        <div style={{ marginTop: 12 }}>
          <h4 style={{ margin: "8px 0" }}>Result</h4>
          <div style={{ fontFamily: "monospace", fontSize: 12 }}>
            <div>Status: {result.status}</div>
            <div>Solution: {result.solution}</div>
            <div>Iterations: {result.iterations}</div>
            <div>Final output: {result.finalOutput}</div>
            <div>Final error: {result.finalError}</div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

