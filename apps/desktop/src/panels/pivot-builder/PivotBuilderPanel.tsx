import React, { useCallback, useMemo, useState } from "react";

import { t, tWithVars } from "../../i18n/index.js";

import type {
  AggregationType,
  PivotField,
  PivotTableConfig,
  ValueField,
} from "./types";

type DropZone = "rows" | "columns" | "values" | "filters";

export interface PivotBuilderPanelProps {
  availableFields: string[];
  initialConfig?: Partial<PivotTableConfig>;
  onChange?: (config: PivotTableConfig) => void;
  onCreate?: (config: PivotTableConfig) => void;
}

const DEFAULT_CONFIG: PivotTableConfig = {
  rowFields: [],
  columnFields: [],
  valueFields: [],
  filterFields: [],
  layout: "tabular",
  subtotals: "none",
  grandTotals: { rows: true, columns: true },
};

function dedupeFields(fields: PivotField[]): PivotField[] {
  const seen = new Set<string>();
  const out: PivotField[] = [];
  for (const f of fields) {
    if (seen.has(f.sourceField)) continue;
    seen.add(f.sourceField);
    out.push(f);
  }
  return out;
}

function defaultValueField(field: string): ValueField {
  return { sourceField: field, name: tWithVars("pivotBuilder.valueField.sumOf", { field }), aggregation: "sum" };
}

export function PivotBuilderPanel({
  availableFields,
  initialConfig,
  onChange,
  onCreate,
}: PivotBuilderPanelProps) {
  const [config, setConfig] = useState<PivotTableConfig>({
    ...DEFAULT_CONFIG,
    ...initialConfig,
    rowFields: initialConfig?.rowFields ?? DEFAULT_CONFIG.rowFields,
    columnFields: initialConfig?.columnFields ?? DEFAULT_CONFIG.columnFields,
    valueFields: initialConfig?.valueFields ?? DEFAULT_CONFIG.valueFields,
    filterFields: initialConfig?.filterFields ?? DEFAULT_CONFIG.filterFields,
    grandTotals: initialConfig?.grandTotals ?? DEFAULT_CONFIG.grandTotals,
  });

  const applyConfig = useCallback(
    (next: PivotTableConfig) => {
      setConfig(next);
      onChange?.(next);
    },
    [onChange]
  );

  const onDragStartField = useCallback((field: string, e: React.DragEvent) => {
    e.dataTransfer.setData("text/plain", field);
    e.dataTransfer.effectAllowed = "copy";
  }, []);

  const onDrop = useCallback(
    (zone: DropZone, e: React.DragEvent) => {
      e.preventDefault();
      const field = e.dataTransfer.getData("text/plain");
      if (!field) return;

      applyConfig({
        ...config,
        rowFields:
          zone === "rows"
            ? dedupeFields([...config.rowFields, { sourceField: field }])
            : config.rowFields,
        columnFields:
          zone === "columns"
            ? dedupeFields([...config.columnFields, { sourceField: field }])
            : config.columnFields,
        valueFields:
          zone === "values"
            ? [
                ...config.valueFields,
                defaultValueField(field),
              ]
            : config.valueFields,
        // Filters are represented by the field name only here; the UI for
        // picking allowed values is left to the worksheet-side integration.
        filterFields:
          zone === "filters"
            ? [
                ...config.filterFields,
                { sourceField: field },
              ]
            : config.filterFields,
      });
    },
    [applyConfig, config]
  );

  const onDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  }, []);

  const aggregations: AggregationType[] = useMemo(
    () => ["sum", "count", "average", "min", "max", "stdDev", "stdDevP", "var", "varP", "countNumbers", "product"],
    []
  );

  const updateValueField = useCallback(
    (idx: number, patch: Partial<ValueField>) => {
      const next = [...config.valueFields];
      next[idx] = { ...next[idx], ...patch };
      applyConfig({ ...config, valueFields: next });
    },
    [applyConfig, config]
  );

  return (
    <div style={{ padding: 12, display: "grid", gridTemplateColumns: "1fr 2fr", gap: 12 }}>
      <section>
        <h3 style={{ margin: "0 0 8px 0" }}>{t("pivotBuilder.fields.title")}</h3>
        <ul style={{ listStyle: "none", padding: 0, margin: 0, display: "grid", gap: 6 }}>
          {availableFields.map((f) => (
            <li
              key={f}
              draggable
              onDragStart={(e) => onDragStartField(f, e)}
              style={{
                padding: "6px 8px",
                border: "1px solid var(--border)",
                borderRadius: 6,
                cursor: "grab",
                userSelect: "none",
              }}
            >
              {f}
            </li>
          ))}
        </ul>
      </section>

      <section style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <DropArea title={t("pivotBuilder.dropArea.rows")} onDrop={(e) => onDrop("rows", e)} onDragOver={onDragOver}>
          {config.rowFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <DropArea title={t("pivotBuilder.dropArea.columns")} onDrop={(e) => onDrop("columns", e)} onDragOver={onDragOver}>
          {config.columnFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <DropArea title={t("pivotBuilder.dropArea.values")} onDrop={(e) => onDrop("values", e)} onDragOver={onDragOver}>
          {config.valueFields.length === 0 ? (
            <div style={{ color: "var(--text-secondary)" }}>{t("pivotBuilder.values.emptyHint")}</div>
          ) : (
            <div style={{ display: "grid", gap: 8 }}>
              {config.valueFields.map((vf, idx) => (
                <div
                  key={`${vf.sourceField}-${idx}`}
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 1fr",
                    gap: 8,
                    padding: 8,
                    border: "1px solid var(--border)",
                    borderRadius: 6,
                  }}
                >
                  <div>
                    <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.value.fieldLabel")}</div>
                    <div>{vf.sourceField}</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 12, color: "var(--text-secondary)" }}>{t("pivotBuilder.value.aggregationLabel")}</div>
                    <select
                      value={vf.aggregation}
                      onChange={(e) => updateValueField(idx, { aggregation: e.target.value as AggregationType })}
                    >
                      {aggregations.map((a) => (
                        <option key={a} value={a}>
                          {t(`pivotBuilder.aggregation.${a}`)}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              ))}
            </div>
          )}
        </DropArea>

        <DropArea title={t("pivotBuilder.dropArea.filters")} onDrop={(e) => onDrop("filters", e)} onDragOver={onDragOver}>
          {config.filterFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <div style={{ gridColumn: "1 / -1", display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button
            onClick={() => applyConfig(DEFAULT_CONFIG)}
            type="button"
            style={{ padding: "6px 10px" }}
          >
            {t("pivotBuilder.actions.reset")}
          </button>
          <button
            onClick={() => onCreate?.(config)}
            type="button"
            disabled={config.valueFields.length === 0}
            style={{ padding: "6px 10px" }}
          >
            {t("pivotBuilder.actions.create")}
          </button>
        </div>
      </section>
    </div>
  );
}

function DropArea({
  title,
  children,
  onDrop,
  onDragOver,
}: {
  title: string;
  children: React.ReactNode;
  onDrop: (e: React.DragEvent) => void;
  onDragOver: (e: React.DragEvent) => void;
}) {
  return (
    <div
      onDrop={onDrop}
      onDragOver={onDragOver}
      style={{
        minHeight: 120,
        padding: 10,
        border: "1px dashed var(--border)",
        borderRadius: 8,
      }}
    >
      <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-secondary)", marginBottom: 8 }}>{title}</div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{children}</div>
    </div>
  );
}

function Pill({ label }: { label: string }) {
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "3px 8px",
        borderRadius: 999,
        background: "var(--bg-tertiary)",
        border: "1px solid var(--border)",
        fontSize: 12,
      }}
    >
      {label}
    </span>
  );
}
