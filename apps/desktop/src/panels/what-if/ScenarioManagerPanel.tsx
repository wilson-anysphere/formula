import React, { useEffect, useMemo, useState } from "react";

import type { Scenario, SummaryReport, WhatIfApi } from "./types";
import { t } from "../../i18n/index.js";

export interface ScenarioManagerPanelProps {
  api: WhatIfApi;
}

export function ScenarioManagerPanel({ api }: ScenarioManagerPanelProps) {
  const [scenarios, setScenarios] = useState<Scenario[]>([]);
  const [selectedScenarioId, setSelectedScenarioId] = useState<number | null>(null);
  const [resultCells, setResultCells] = useState("B1");
  const [report, setReport] = useState<SummaryReport | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void (async () => {
      try {
        setError(null);
        const list = await api.listScenarios();
        setScenarios(list);
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      }
    })();
  }, [api]);

  const selectedScenario = useMemo(
    () => scenarios.find((s) => s.id === selectedScenarioId) ?? null,
    [scenarios, selectedScenarioId]
  );

  async function applySelected() {
    if (selectedScenarioId == null) return;
    setError(null);
    setReport(null);
    try {
      await api.applyScenario(selectedScenarioId);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  async function restoreBase() {
    setError(null);
    setReport(null);
    try {
      await api.restoreBaseScenario();
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  async function generateReport() {
    setError(null);
    setReport(null);

    const cells = resultCells
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);

    if (cells.length === 0) {
      setError(t("whatIf.scenario.error.enterResultCell"));
      return;
    }

    try {
      const ids = scenarios.map((s) => s.id);
      const summary = await api.generateSummaryReport(cells, ids);
      setReport(summary);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    }
  }

  return (
    <div style={{ padding: 16, border: "1px solid #ccc", borderRadius: 8 }}>
      <h3 style={{ marginTop: 0 }}>{t("whatIf.scenario.title")}</h3>

      {error ? <p style={{ color: "crimson" }}>{error}</p> : null}

      <div style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "center" }}>
        <select
          value={selectedScenarioId ?? ""}
          onChange={(e) => setSelectedScenarioId(e.target.value ? Number(e.target.value) : null)}
        >
          <option value="">{t("whatIf.scenario.selectPlaceholder")}</option>
          {scenarios.map((s) => (
            <option key={s.id} value={s.id}>
              {s.name}
            </option>
          ))}
        </select>

        <button onClick={applySelected} disabled={selectedScenarioId == null}>
          {t("whatIf.scenario.apply")}
        </button>
      </div>

      {selectedScenario ? (
        <div style={{ marginTop: 8, fontSize: 12, color: "#444" }}>
          <div>
            {t("whatIf.scenario.changingCells")}: {selectedScenario.changingCells.join(", ") || "—"}
          </div>
          {selectedScenario.comment ? (
            <div>
              {t("whatIf.scenario.comment")}: {selectedScenario.comment}
            </div>
          ) : null}
        </div>
      ) : null}

      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <button onClick={restoreBase}>{t("whatIf.scenario.restoreBase")}</button>
        <button onClick={generateReport} disabled={scenarios.length === 0}>
          {t("whatIf.scenario.summaryReport")}
        </button>
      </div>

      <div style={{ marginTop: 12 }}>
        <label style={{ display: "grid", gap: 4 }}>
          <span>{t("whatIf.scenario.resultCellsLabel")}</span>
          <input value={resultCells} onChange={(e) => setResultCells(e.target.value)} />
        </label>
      </div>

      {report ? (
        <div style={{ marginTop: 16 }}>
          <h4 style={{ margin: "8px 0" }}>{t("whatIf.scenario.summaryTitle")}</h4>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
                  {t("whatIf.scenario.table.scenario")}
                </th>
                {report.resultCells.map((cell) => (
                  <th key={cell} style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
                    {cell}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Object.entries(report.results).map(([scenarioName, row]) => (
                <tr key={scenarioName}>
                  <td style={{ padding: "4px 0" }}>{scenarioName}</td>
                  {report.resultCells.map((cell) => (
                    <td key={cell} style={{ padding: "4px 0" }}>
                      {String(row[cell] ?? "")}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : null}
    </div>
  );
}
