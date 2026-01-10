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

function defaultDistribution(): InputDistribution {
  const dist: Distribution = { type: "normal", mean: 0, stdDev: 1 };
  return { cell: "A1", distribution: dist };
}

export function MonteCarloWizard({ api }: MonteCarloWizardProps) {
  const [iterations, setIterations] = useState("1000");
  const [seed, setSeed] = useState("1234");
  const [outputCells, setOutputCells] = useState("B1");

  const [inputs, setInputs] = useState<InputDistribution[]>([defaultDistribution()]);
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState<SimulationProgress | null>(null);
  const [result, setResult] = useState<SimulationResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const parsedIterations = useMemo(() => Number(iterations), [iterations]);
  const parsedSeed = useMemo(() => Number(seed), [seed]);

  function updateInput(idx: number, patch: Partial<InputDistribution>) {
    setInputs((prev) => prev.map((v, i) => (i === idx ? { ...v, ...patch } : v)));
  }

  async function run() {
    setError(null);
    setResult(null);
    setProgress(null);

    if (!Number.isFinite(parsedIterations) || parsedIterations <= 0) {
      setError(t("whatIf.monteCarlo.error.iterationsPositive"));
      return;
    }

    const outputs = outputCells
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);

    if (outputs.length === 0) {
      setError(t("whatIf.monteCarlo.error.enterOutputCell"));
      return;
    }

    const config: SimulationConfig = {
      iterations: Math.floor(parsedIterations),
      seed: Number.isFinite(parsedSeed) ? Math.floor(parsedSeed) : undefined,
      inputDistributions: inputs.map((i) => ({
        ...i,
        cell: i.cell.trim(),
      })),
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
    <div style={{ padding: 16, border: "1px solid var(--panel-border)", borderRadius: 8 }}>
      <h3 style={{ marginTop: 0 }}>{t("whatIf.monteCarlo.title")}</h3>

      {error ? <p style={{ color: "var(--error)" }}>{error}</p> : null}

      <div style={{ display: "grid", gap: 8, gridTemplateColumns: "1fr 1fr 1fr", alignItems: "end" }}>
        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.monteCarlo.iterations")}</span>
          <input value={iterations} onChange={(e) => setIterations(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.monteCarlo.seed")}</span>
          <input value={seed} onChange={(e) => setSeed(e.target.value)} disabled={running} />
        </label>

        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.monteCarlo.outputCells")}</span>
          <input value={outputCells} onChange={(e) => setOutputCells(e.target.value)} disabled={running} />
        </label>
      </div>

      <div style={{ marginTop: 16 }}>
        <h4 style={{ margin: "8px 0" }}>{t("whatIf.monteCarlo.inputs")}</h4>
        <div style={{ display: "grid", gap: 8 }}>
          {inputs.map((input, idx) => (
            <div
              key={idx}
              style={{
                display: "grid",
                gridTemplateColumns: "110px 1fr 1fr auto",
                gap: 8,
                alignItems: "center",
              }}
            >
              <input
                value={input.cell}
                onChange={(e) => updateInput(idx, { cell: e.target.value })}
                disabled={running}
                placeholder="A1"
              />

              <select
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
                  updateInput(idx, { distribution });
                }}
                disabled={running}
              >
                <option value="normal">{t("whatIf.distribution.normal")}</option>
                <option value="uniform">{t("whatIf.distribution.uniform")}</option>
                <option value="triangular">{t("whatIf.distribution.triangular")}</option>
                <option value="lognormal">{t("whatIf.distribution.lognormal")}</option>
                <option value="exponential">{t("whatIf.distribution.exponential")}</option>
                <option value="poisson">{t("whatIf.distribution.poisson")}</option>
              </select>

              <input
                value={JSON.stringify(input.distribution)}
                onChange={(e) => {
                  try {
                    const parsed = JSON.parse(e.target.value) as Distribution;
                    updateInput(idx, { distribution: parsed });
                  } catch {
                    // Allow partial JSON edits.
                  }
                }}
                disabled={running}
                style={{ fontFamily: "monospace", fontSize: 12 }}
              />

              <button
                onClick={() => setInputs((prev) => prev.filter((_, i) => i !== idx))}
                disabled={running || inputs.length <= 1}
              >
                {t("whatIf.monteCarlo.remove")}
              </button>
            </div>
          ))}

          <div>
            <button onClick={() => setInputs((prev) => [...prev, defaultDistribution()])} disabled={running}>
              {t("whatIf.monteCarlo.addInput")}
            </button>
          </div>
        </div>
      </div>

      <div style={{ marginTop: 16 }}>
        <button onClick={run} disabled={running}>
          {running ? t("whatIf.monteCarlo.running") : t("whatIf.monteCarlo.runSimulation")}
        </button>
      </div>

      {progress ? (
        <p style={{ marginTop: 12, fontFamily: "monospace", fontSize: 12 }}>
          {tWithVars("whatIf.monteCarlo.progressIterations", {
            completed: progress.completedIterations,
            total: progress.totalIterations,
          })}
        </p>
      ) : null}

      {result ? (
        <div style={{ marginTop: 16 }}>
          <h4 style={{ margin: "8px 0" }}>{t("whatIf.monteCarlo.results")}</h4>
          {Object.entries(result.outputStats).map(([cell, stats]) => (
            <div key={cell} style={{ marginBottom: 12 }}>
              <strong>{cell}</strong>
              <div style={{ fontFamily: "monospace", fontSize: 12 }}>
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
      ) : null}
    </div>
  );
}
