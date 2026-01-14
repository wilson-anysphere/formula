import React, { useEffect, useMemo, useState } from "react";

import type { ArrowTableAdapter, DataTable, Query, QueryEngine, QueryOperation, QueryStep } from "@formula/power-query";

import { StepsList } from "./components/StepsList";
import { PreviewGrid } from "./components/PreviewGrid";
import { SchemaView } from "./components/SchemaView";
import { AddStepMenu } from "./components/AddStepMenu";
import { formatQueryOperationLabel } from "./operationLabels";

export type QueryEditorPanelProps = {
  query: Query;
  engine: QueryEngine;
  context?: any;
  onQueryChange?: (next: Query) => void;
  onLoadToSheet?: (query: Query) => void;
  onRefreshNow?: (queryId: string) => void;
  refreshEvent?: unknown;
  onAiSuggestNextSteps?: (
    intent: string,
    context: { query: Query; preview: DataTable | ArrowTableAdapter | null },
  ) => Promise<QueryOperation[]>;
};

/**
 * Query editor panel – "Power Query" equivalent UX.
 *
 * This is a thin UI wrapper around the `packages/power-query` execution engine.
 * It intentionally renders a limited preview (first 100 rows) to keep the UI
 * responsive, even for large sources.
 */
export function QueryEditorPanel(props: QueryEditorPanelProps) {
  const [selectedStepIndex, setSelectedStepIndex] = useState<number>(props.query.steps.length - 1);
  const [preview, setPreview] = useState<DataTable | ArrowTableAdapter | null>(null);
  const [error, setError] = useState<string | null>(null);

  const effectiveSelectedStepIndex = Math.max(-1, Math.min(selectedStepIndex, props.query.steps.length - 1));
  const stepsToExecute = useMemo<QueryStep[]>(
    () => (effectiveSelectedStepIndex < 0 ? [] : props.query.steps.slice(0, effectiveSelectedStepIndex + 1)),
    [props.query.steps, effectiveSelectedStepIndex],
  );
  const queryAtSelectedStep = useMemo<Query>(() => ({ ...props.query, steps: stepsToExecute }), [props.query, stepsToExecute]);

  useEffect(() => {
    let cancelled = false;
    const controller = new AbortController();
    (async () => {
      try {
        setError(null);
        const result = await props.engine.executeQuery(
          { ...props.query, steps: stepsToExecute },
          props.context ?? {},
          { limit: 100, signal: controller.signal },
        );
        if (!cancelled) setPreview(result);
      } catch (e: any) {
        if (cancelled) return;
        if (e?.name === "AbortError") return;
        setError(e?.message ?? String(e));
      }
    })();
    return () => {
      cancelled = true;
      controller.abort();
    };
  }, [props.engine, props.query, props.context, stepsToExecute]);

  const refreshStatus = useMemo(() => {
    const evt: any = props.refreshEvent;
    if (!evt) return null;
    const jobQueryId = evt?.job?.queryId;
    const applyQueryId = evt?.queryId;
    if (typeof jobQueryId === "string" && jobQueryId !== props.query.id) return null;
    if (typeof jobQueryId !== "string" && typeof applyQueryId === "string" && applyQueryId !== props.query.id) return null;
    switch (evt.type) {
      case "queued":
        return "Refresh queued…";
      case "started":
        return "Refreshing…";
      case "progress":
        return `Refreshing… (${evt?.event?.type ?? "working"})`;
      case "completed":
        return "Refresh complete";
      case "cancelled":
        return "Refresh cancelled";
      case "error":
        return `Refresh failed: ${evt?.error?.message ?? String(evt?.error ?? "Unknown error")}`;
      case "apply:started":
        return "Applying results to sheet…";
      case "apply:progress":
        return `Applying results… (${evt?.rowsWritten ?? 0} rows)`;
      case "apply:completed":
        return "Results applied";
      case "apply:cancelled":
        return "Apply cancelled";
      case "apply:error":
        return `Apply failed: ${evt?.error?.message ?? String(evt?.error ?? "Unknown error")}`;
      default:
        return null;
    }
  }, [props.refreshEvent, props.query.id]);

  return (
    <div className="query-editor">
      <div className="query-editor__sidebar">
        <h3 className="query-editor__title">{props.query.name}</h3>
        <div className="query-editor__sidebar-actions">
          {props.onLoadToSheet ? (
            <button type="button" onClick={() => props.onLoadToSheet?.(props.query)}>
              Load to sheet
            </button>
          ) : null}
          {props.onRefreshNow ? (
            <button type="button" onClick={() => props.onRefreshNow?.(props.query.id)}>
              Refresh now
            </button>
          ) : null}
        </div>
        <AddStepMenu
          onAddStep={(operation) => {
            const baseName = formatQueryOperationLabel(operation);
            const existingNames = new Set(props.query.steps.map((step) => step.name));
            let name = baseName;
            let suffix = 1;
            while (existingNames.has(name)) {
              name = `${baseName} ${suffix}`;
              suffix += 1;
            }

            const step: QueryStep = {
              id: crypto.randomUUID(),
              name,
              operation,
            };
            const insertAt = effectiveSelectedStepIndex + 1;
            const nextSteps = props.query.steps.slice();
            nextSteps.splice(insertAt, 0, step);
            props.onQueryChange?.({ ...props.query, steps: nextSteps });
            setSelectedStepIndex(insertAt);
          }}
          onAiSuggest={props.onAiSuggestNextSteps}
          aiContext={{ query: queryAtSelectedStep, preview }}
        />
        <StepsList steps={props.query.steps} selectedIndex={effectiveSelectedStepIndex} onSelect={setSelectedStepIndex} />
      </div>

      <div className="query-editor__main">
        <div className="query-editor__schema">
          <SchemaView table={preview} />
          {refreshStatus ? <div className="query-editor__status">{refreshStatus}</div> : null}
          {error ? <div className="query-editor__error">{error}</div> : null}
        </div>
        <div className="query-editor__preview">
          <PreviewGrid table={preview} />
        </div>
      </div>
    </div>
  );
}
