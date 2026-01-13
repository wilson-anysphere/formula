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
  createDisabled?: boolean;
  createLabel?: string;
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

function testIdSafe(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]+/g, "-");
}

export function PivotBuilderPanel({
  availableFields,
  initialConfig,
  onChange,
  onCreate,
  createDisabled,
  createLabel,
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
    <div className="pivot-builder__builder">
      <section>
        <h3 className="pivot-builder__fields-title">{t("pivotBuilder.fields.title")}</h3>
        <ul className="pivot-builder__fields-list">
          {availableFields.map((f) => (
            <li
              key={f}
              data-testid={`pivot-field-${testIdSafe(f)}`}
              draggable
              onDragStart={(e) => onDragStartField(f, e)}
              className="pivot-builder__field"
            >
              {f}
            </li>
          ))}
        </ul>
      </section>

      <section className="pivot-builder__drop-zones">
        <DropArea
          title={t("pivotBuilder.dropArea.rows")}
          testId="pivot-drop-rows"
          onDrop={(e) => onDrop("rows", e)}
          onDragOver={onDragOver}
        >
          {config.rowFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <DropArea
          title={t("pivotBuilder.dropArea.columns")}
          testId="pivot-drop-columns"
          onDrop={(e) => onDrop("columns", e)}
          onDragOver={onDragOver}
        >
          {config.columnFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <DropArea
          title={t("pivotBuilder.dropArea.values")}
          testId="pivot-drop-values"
          onDrop={(e) => onDrop("values", e)}
          onDragOver={onDragOver}
        >
          {config.valueFields.length === 0 ? (
            <div className="pivot-builder__empty-hint">{t("pivotBuilder.values.emptyHint")}</div>
          ) : (
            <div className="pivot-builder__value-fields">
              {config.valueFields.map((vf, idx) => (
                <div
                  key={`${vf.sourceField}-${idx}`}
                  className="pivot-builder__value-field"
                >
                  <div>
                    <div className="pivot-builder__meta-label">{t("pivotBuilder.value.fieldLabel")}</div>
                    <div>{vf.sourceField}</div>
                  </div>
                  <div>
                    <div className="pivot-builder__meta-label">{t("pivotBuilder.value.aggregationLabel")}</div>
                    <select
                      data-testid={`pivot-value-aggregation-${idx}`}
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

        <DropArea
          title={t("pivotBuilder.dropArea.filters")}
          testId="pivot-drop-filters"
          onDrop={(e) => onDrop("filters", e)}
          onDragOver={onDragOver}
        >
          {config.filterFields.map((f) => (
            <Pill key={f.sourceField} label={f.sourceField} />
          ))}
        </DropArea>

        <div className="pivot-builder__options-row">
          <label className="pivot-builder__label">
            <input
              data-testid="pivot-grand-totals-rows"
              type="checkbox"
              checked={config.grandTotals.rows}
              onChange={(e) => applyConfig({ ...config, grandTotals: { ...config.grandTotals, rows: e.target.checked } })}
            />
            {t("pivotBuilder.options.grandTotalsRows")}
          </label>
          <label className="pivot-builder__label">
            <input
              data-testid="pivot-grand-totals-columns"
              type="checkbox"
              checked={config.grandTotals.columns}
              onChange={(e) =>
                applyConfig({ ...config, grandTotals: { ...config.grandTotals, columns: e.target.checked } })
              }
            />
            {t("pivotBuilder.options.grandTotalsColumns")}
          </label>
        </div>

        <div className="pivot-builder__actions-row">
          <button onClick={() => applyConfig(DEFAULT_CONFIG)} type="button" data-testid="pivot-reset">
            {t("pivotBuilder.actions.reset")}
          </button>
          <button
            onClick={() => onCreate?.(config)}
            type="button"
            disabled={Boolean(createDisabled) || config.valueFields.length === 0}
            data-testid="pivot-create"
          >
            {createLabel ?? t("pivotBuilder.actions.create")}
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
  testId,
}: {
  title: string;
  children: React.ReactNode;
  onDrop: (e: React.DragEvent) => void;
  onDragOver: (e: React.DragEvent) => void;
  testId?: string;
}) {
  return (
    <div
      data-testid={testId}
      onDrop={onDrop}
      onDragOver={onDragOver}
      className="pivot-builder__drop-area"
    >
      <div className="pivot-builder__drop-title">{title}</div>
      <div className="pivot-builder__drop-content">{children}</div>
    </div>
  );
}

function Pill({ label }: { label: string }) {
  return (
    <span className="pivot-builder__pill">{label}</span>
  );
}
