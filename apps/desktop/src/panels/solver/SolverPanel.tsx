import React, { useCallback, useMemo, useRef, useState } from "react";

import { t } from "../../i18n/index.js";

import { runSolver } from "./api";
import { SolverDialog } from "./SolverDialog";
import { SolverProgressView } from "./SolverProgress";
import { SolverResultSummary } from "./SolverResultSummary";
import type { SolverConfig, SolverOutcome, SolverProgress } from "./types";

export function SolverPanel() {
  const [showDialog, setShowDialog] = useState(false);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<SolverProgress | null>(null);
  const [outcome, setOutcome] = useState<SolverOutcome | null>(null);
  const [error, setError] = useState<string | null>(null);

  const abortRef = useRef<AbortController | null>(null);

  const onRun = useCallback(async (config: SolverConfig) => {
    setShowDialog(false);
    setRunning(true);
    setError(null);
    setOutcome(null);
    setProgress(null);

    const abort = new AbortController();
    abortRef.current = abort;

    try {
      const res = await runSolver(
        config,
        (p) => setProgress(p),
        abort.signal,
      );
      setOutcome(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setRunning(false);
      abortRef.current = null;
    }
  }, []);

  const onCancelRun = useCallback(() => {
    abortRef.current?.abort();
  }, []);

  const onKeep = useCallback(() => {
    // In a full implementation this would commit the solver solution (already applied by the engine)
    // and close the summary.
    setOutcome(null);
  }, []);

  const onRestore = useCallback(() => {
    // In a full implementation this would restore the original variable values.
    setOutcome(null);
  }, []);

  const dialogInitial = useMemo(() => {
    // In the real application, this should be derived from the current sheet selection.
    return undefined;
  }, []);

  return (
    <div className="solver-panel">
      <header className="solver-panel__header">
        <h2 className="solver-panel__title">{t("panels.solver.title")}</h2>
        <button
          type="button"
          className="solver-panel__configure-button"
          onClick={() => setShowDialog(true)}
        >
          {t("solver.configure")}
        </button>
      </header>

      {showDialog && (
        <SolverDialog
          initial={dialogInitial}
          onCancel={() => setShowDialog(false)}
          onRun={onRun}
        />
      )}

      {running && (
        <SolverProgressView progress={progress} onCancel={onCancelRun} />
      )}

      {outcome && (
        <SolverResultSummary outcome={outcome} onKeep={onKeep} onRestore={onRestore} />
      )}

      {error && (
        <div className="solver-panel__error">
          <strong>{t("solver.errorLabel")}</strong> {error}
        </div>
      )}

      {!running && !outcome && !error && !showDialog && (
        <p>{t("solver.emptyState")}</p>
      )}
    </div>
  );
}
