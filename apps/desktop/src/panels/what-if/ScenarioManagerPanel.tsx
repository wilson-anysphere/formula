import React, { useEffect, useMemo, useState } from "react";

import type { Scenario, SummaryReport, WhatIfApi, WhatIfCellValue } from "./types";
import { t } from "../../i18n/index.js";

export interface ScenarioManagerPanelProps {
  api: WhatIfApi;
}

function formatWhatIfCellValue(value: WhatIfCellValue | undefined): string {
  if (!value) return "";
  switch (value.type) {
    case "number":
      return String(value.value);
    case "text":
      return value.value;
    case "bool":
      return value.value ? "TRUE" : "FALSE";
    case "blank":
      return "";
    default:
      // Exhaustive check: ensure we still render something if the backend adds a new type.
      return String((value as unknown) ?? "");
  }
}

export function ScenarioManagerPanel({ api }: ScenarioManagerPanelProps) {
  const [scenarios, setScenarios] = useState<Scenario[]>([]);
  const [selectedScenarioId, setSelectedScenarioId] = useState<number | null>(null);
  const [resultCells, setResultCells] = useState("B1");
  const [report, setReport] = useState<SummaryReport | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [invalidField, setInvalidField] = useState<"resultCells" | null>(null);
  const [busy, setBusy] = useState(false);
  const reactInstanceId = React.useId();
  const domInstanceId = useMemo(() => reactInstanceId.replace(/[^a-zA-Z0-9_-]/g, "-"), [reactInstanceId]);
  const errorId = useMemo(() => `scenario-manager-error-${domInstanceId}`, [domInstanceId]);

  useEffect(() => {
    void (async () => {
      try {
        setBusy(true);
        setError(null);
        const list = await api.listScenarios();
        setScenarios(list);
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setBusy(false);
      }
    })().catch(() => {});
  }, [api]);

  const selectedScenario = useMemo(
    () => scenarios.find((s) => s.id === selectedScenarioId) ?? null,
    [scenarios, selectedScenarioId]
  );

  async function applySelected() {
    if (selectedScenarioId == null) return;
    setBusy(true);
    setError(null);
    setReport(null);
    try {
      await api.applyScenario(selectedScenarioId);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function restoreBase() {
    setBusy(true);
    setError(null);
    setReport(null);
    try {
      await api.restoreBaseScenario();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function generateReport() {
    setError(null);
    setInvalidField(null);
    setReport(null);

    const cells = resultCells
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);

    if (cells.length === 0) {
      setError(t("whatIf.scenario.error.enterResultCell"));
      setInvalidField("resultCells");
      return;
    }

    setBusy(true);
    try {
      const ids = scenarios.map((s) => s.id);
      const summary = await api.generateSummaryReport(cells, ids);
      setReport(summary);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div
      className="what-if-panel"
      role="region"
      aria-label={t("whatIf.scenario.title")}
      aria-busy={busy ? true : undefined}
      data-testid="scenario-manager-panel"
    >
      <h3 className="what-if-panel__title">{t("whatIf.scenario.title")}</h3>

      {error ? (
        <p className="what-if__message what-if__message--error" role="alert" id={errorId}>
          {error}
        </p>
      ) : null}

      <div className="what-if__row">
        <label className="what-if__field what-if__field--grow">
          <span className="what-if__label">{t("whatIf.scenario.table.scenario")}</span>
          <select
            className="what-if__select"
            value={selectedScenarioId ?? ""}
            onChange={(e) => setSelectedScenarioId(e.target.value ? Number(e.target.value) : null)}
            disabled={busy}
          >
            <option value="">{t("whatIf.scenario.selectPlaceholder")}</option>
            {scenarios.map((s) => (
              <option key={s.id} value={s.id}>
                {s.name}
              </option>
            ))}
          </select>
        </label>

        <button
          type="button"
          className="what-if__button what-if__button--primary"
          onClick={applySelected}
          disabled={busy || selectedScenarioId == null}
        >
          {t("whatIf.scenario.apply")}
        </button>
      </div>

      {selectedScenario ? (
        <div className="what-if__meta" data-testid="scenario-manager-selected-meta">
          <div className="what-if__meta-row">
            <span className="what-if__meta-label">{t("whatIf.scenario.changingCells")}:</span>
            <span className="what-if__meta-value">{selectedScenario.changingCells.join(", ") || "—"}</span>
          </div>
          {selectedScenario.comment ? (
            <div className="what-if__meta-row">
              <span className="what-if__meta-label">{t("whatIf.scenario.comment")}:</span>
              <span className="what-if__meta-value">{selectedScenario.comment}</span>
            </div>
          ) : null}
        </div>
      ) : null}

      <div className="what-if__actions">
        <button type="button" className="what-if__button" onClick={restoreBase} disabled={busy}>
          {t("whatIf.scenario.restoreBase")}
        </button>
        <button type="button" className="what-if__button" onClick={generateReport} disabled={busy || scenarios.length === 0}>
          {t("whatIf.scenario.summaryReport")}
        </button>
      </div>

      <label className="what-if__field">
        <span className="what-if__label">{t("whatIf.scenario.resultCellsLabel")}</span>
        <input
          className="what-if__input what-if__input--mono"
          value={resultCells}
          onChange={(e) => {
            setResultCells(e.target.value);
            if (invalidField === "resultCells") {
              setInvalidField(null);
              setError(null);
            }
          }}
          spellCheck={false}
          autoCapitalize="off"
          aria-invalid={invalidField === "resultCells" ? true : undefined}
          disabled={busy}
          aria-describedby={invalidField === "resultCells" && error ? errorId : undefined}
        />
      </label>

      {report ? (
        <div className="what-if__section" data-testid="scenario-manager-report">
          <h4 className="what-if__section-title">{t("whatIf.scenario.summaryTitle")}</h4>
          <div className="what-if__table-wrap">
            <table className="what-if-table" aria-label={t("whatIf.scenario.summaryTitle")}>
              <thead>
                <tr>
                  <th scope="col">{t("whatIf.scenario.table.scenario")}</th>
                  {report.resultCells.map((cell) => (
                    <th scope="col" key={cell}>
                      {cell}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(report.results).map(([scenarioName, row]) => (
                  <tr key={scenarioName}>
                    <th scope="row">{scenarioName}</th>
                    {report.resultCells.map((cell) => (
                      <td key={cell}>{formatWhatIfCellValue(row[cell])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : null}
    </div>
  );
}
