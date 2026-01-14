import React, { useMemo, useState } from "react";

import type {
  Distribution,
  InputDistribution,
  SimulationConfig,
  SimulationProgress,
  SimulationResult,
  WhatIfApi,
} from "./types";
import { t, tWithVars } from "../../i18n/index.js";

export interface MonteCarloWizardProps {
  api: WhatIfApi;
}

type InputRow = InputDistribution & { distributionJson: string };

function defaultInputRow(): InputRow {
  const dist: Distribution = { type: "normal", mean: 0, stdDev: 1 };
  return { cell: "A1", distribution: dist, distributionJson: JSON.stringify(dist) };
}

export function MonteCarloWizard({ api }: MonteCarloWizardProps) {
  const [iterations, setIterations] = useState("1000");
  const [seed, setSeed] = useState("1234");
  const [outputCells, setOutputCells] = useState("B1");

  const [inputs, setInputs] = useState<InputRow[]>([defaultInputRow()]);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<SimulationProgress | null>(null);
  const [result, setResult] = useState<SimulationResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [invalidField, setInvalidField] = useState<"iterations" | "outputCells" | null>(null);
  const reactInstanceId = React.useId();
  const domInstanceId = useMemo(() => reactInstanceId.replace(/[^a-zA-Z0-9_-]/g, "-"), [reactInstanceId]);
  const errorId = useMemo(() => `monte-carlo-error-${domInstanceId}`, [domInstanceId]);

  const parsedIterations = useMemo(() => Number(iterations), [iterations]);
  const parsedSeed = useMemo(() => Number(seed), [seed]);

  function updateInput(idx: number, patch: Partial<InputRow>) {
    setInputs((prev) => prev.map((v, i) => (i === idx ? { ...v, ...patch } : v)));
  }

  async function run() {
    setError(null);
    setInvalidField(null);
    setResult(null);
    setProgress(null);

    if (!Number.isFinite(parsedIterations) || parsedIterations <= 0) {
      setError(t("whatIf.monteCarlo.error.iterationsPositive"));
      setInvalidField("iterations");
      return;
    }

    const outputs = outputCells
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);

    if (outputs.length === 0) {
      setError(t("whatIf.monteCarlo.error.enterOutputCell"));
      setInvalidField("outputCells");
      return;
    }

    const config: SimulationConfig = {
      iterations: Math.floor(parsedIterations),
      seed: Number.isFinite(parsedSeed) ? Math.floor(parsedSeed) : undefined,
      inputDistributions: inputs.map(({ cell, distribution }) => ({ cell: cell.trim(), distribution })),
      outputCells: outputs,
    };

    setRunning(true);
    try {
      const res = await api.runMonteCarlo(config, (p) => setProgress(p));
      setResult(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setRunning(false);
    }
  }

  return (
    <div
      className="what-if-panel"
      role="region"
      aria-label={t("whatIf.monteCarlo.title")}
      aria-busy={running ? true : undefined}
      data-testid="monte-carlo-wizard"
    >
      <h3 className="what-if-panel__title">{t("whatIf.monteCarlo.title")}</h3>

      {error ? (
        <p className="what-if__message what-if__message--error" role="alert" id={errorId}>
          {error}
        </p>
      ) : null}

      <div className="what-if-grid what-if-grid--fit">
        <label className="what-if__field">
          <span className="what-if__label">{t("whatIf.monteCarlo.iterations")}</span>
          <input
            className="what-if__input"
            value={iterations}
            onChange={(e) => {
              setIterations(e.target.value);
              if (invalidField === "iterations") {
                setInvalidField(null);
                setError(null);
              }
            }}
            disabled={running}
            inputMode="numeric"
            aria-invalid={invalidField === "iterations" ? true : undefined}
            aria-describedby={invalidField === "iterations" && error ? errorId : undefined}
          />
        </label>

        <label className="what-if__field">
          <span className="what-if__label">{t("whatIf.monteCarlo.seed")}</span>
          <input className="what-if__input" value={seed} onChange={(e) => setSeed(e.target.value)} disabled={running} inputMode="numeric" />
        </label>

        <label className="what-if__field">
          <span className="what-if__label">{t("whatIf.monteCarlo.outputCells")}</span>
          <input
            className="what-if__input what-if__input--mono"
            value={outputCells}
            onChange={(e) => {
              setOutputCells(e.target.value);
              if (invalidField === "outputCells") {
                setInvalidField(null);
                setError(null);
              }
            }}
            disabled={running}
            spellCheck={false}
            autoCapitalize="off"
            aria-invalid={invalidField === "outputCells" ? true : undefined}
            aria-describedby={invalidField === "outputCells" && error ? errorId : undefined}
          />
        </label>
      </div>

      <div className="what-if__section">
        <h4 className="what-if__section-title">{t("whatIf.monteCarlo.inputs")}</h4>
        <div className="what-if-monte-carlo__inputs">
          {inputs.map((input, idx) => (
            <div key={idx} className="what-if-monte-carlo__input-row" data-testid={`monte-carlo-input-${idx}`}>
              <div className="what-if-monte-carlo__input-cell">
                <input
                  className="what-if__input what-if__input--mono"
                  value={input.cell}
                  onChange={(e) => updateInput(idx, { cell: e.target.value })}
                  disabled={running}
                  placeholder="A1"
                  aria-label={t("whatIf.monteCarlo.inputs.cellAriaLabel")}
                  spellCheck={false}
                  autoCapitalize="off"
                />
              </div>

              <div className="what-if-monte-carlo__input-type">
                <select
                  className="what-if__select"
                  value={input.distribution.type}
                  onChange={(e) => {
                    const type = e.target.value as Distribution["type"];
                    // Keep it simple: switching resets to a reasonable default.
                    const distribution: Distribution =
                      type === "normal"
                        ? { type, mean: 0, stdDev: 1 }
                        : type === "uniform"
                          ? { type, min: 0, max: 1 }
                          : type === "triangular"
                            ? { type, min: 0, mode: 0.5, max: 1 }
                            : type === "lognormal"
                              ? { type, mean: 0, stdDev: 1 }
                              : type === "exponential"
                                ? { type, rate: 1 }
                                : type === "poisson"
                                  ? { type, lambda: 1 }
                                  : { type: "normal", mean: 0, stdDev: 1 };
                    updateInput(idx, { distribution, distributionJson: JSON.stringify(distribution) });
                  }}
                  disabled={running}
                  aria-label={t("whatIf.monteCarlo.inputs.distributionTypeAriaLabel")}
                >
                  <option value="normal">{t("whatIf.distribution.normal")}</option>
                  <option value="uniform">{t("whatIf.distribution.uniform")}</option>
                  <option value="triangular">{t("whatIf.distribution.triangular")}</option>
                  <option value="lognormal">{t("whatIf.distribution.lognormal")}</option>
                  <option value="exponential">{t("whatIf.distribution.exponential")}</option>
                  <option value="poisson">{t("whatIf.distribution.poisson")}</option>
                </select>
              </div>

              <div className="what-if-monte-carlo__input-json">
                <input
                  className="what-if__input what-if__input--mono"
                  value={input.distributionJson}
                  onChange={(e) => {
                    const raw = e.target.value;
                    try {
                      const parsed = JSON.parse(raw) as unknown;
                      if (parsed && typeof parsed === "object" && typeof (parsed as any).type === "string") {
                        updateInput(idx, { distribution: parsed as Distribution, distributionJson: raw });
                        return;
                      }
                    } catch {
                      // Allow partial JSON edits.
                    }
                    updateInput(idx, { distributionJson: raw });
                  }}
                  disabled={running}
                  aria-label={t("whatIf.monteCarlo.inputs.distributionJsonAriaLabel")}
                  spellCheck={false}
                  autoCapitalize="off"
                />
              </div>

              <div className="what-if-monte-carlo__input-actions">
                <button
                  type="button"
                  className="what-if__button"
                  onClick={() => setInputs((prev) => prev.filter((_, i) => i !== idx))}
                  disabled={running || inputs.length <= 1}
                >
                  {t("whatIf.monteCarlo.remove")}
                </button>
              </div>
            </div>
          ))}

          <div className="what-if__actions">
            <button
              type="button"
              className="what-if__button"
              onClick={() => setInputs((prev) => [...prev, defaultInputRow()])}
              disabled={running}
            >
              {t("whatIf.monteCarlo.addInput")}
            </button>
          </div>
        </div>
      </div>

      <div className="what-if__actions">
        <button type="button" className="what-if__button what-if__button--primary" onClick={run} disabled={running}>
          {running ? t("whatIf.monteCarlo.running") : t("whatIf.monteCarlo.runSimulation")}
        </button>
      </div>

      {progress ? (
        <div className="what-if__message what-if__mono-block" role="status" data-testid="monte-carlo-progress">
          {tWithVars("whatIf.monteCarlo.progressIterations", {
            completed: progress.completedIterations,
            total: progress.totalIterations,
          })}
        </div>
      ) : null}

      {result ? (
        <div className="what-if__section" data-testid="monte-carlo-results">
          <h4 className="what-if__section-title">{t("whatIf.monteCarlo.results")}</h4>
          <div className="what-if-grid">
            {Object.entries(result.outputStats).map(([cell, stats]) => (
              <div key={cell} className="what-if-monte-carlo__result">
                <div className="what-if-monte-carlo__result-title">{cell}</div>
                <div className="what-if__mono-block">
                  <div>
                    {t("whatIf.stats.mean")}: {stats.mean}
                  </div>
                  <div>
                    {t("whatIf.stats.median")}: {stats.median}
                  </div>
                  <div>
                    {t("whatIf.stats.stdDev")}: {stats.stdDev}
                  </div>
                  <div>
                    {t("whatIf.stats.minMax")}: {stats.min} / {stats.max}
                  </div>
                  <div>
                    {tWithVars("whatIf.stats.percentile", { p: 5 })}: {stats.percentiles["5"]}
                  </div>
                  <div>
                    {tWithVars("whatIf.stats.percentile", { p: 95 })}: {stats.percentiles["95"]}
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
    </div>
  );
}
