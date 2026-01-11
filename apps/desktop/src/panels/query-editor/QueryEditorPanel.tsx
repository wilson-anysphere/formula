import React, { useEffect, useMemo, useState } from "react";

import type { Query, QueryOperation, QueryStep } from "../../../../packages/power-query/src/model.js";
import { QueryEngine } from "../../../../packages/power-query/src/engine.js";
import type { DataTable } from "../../../../packages/power-query/src/table.js";
import type { ArrowTableAdapter } from "../../../../packages/power-query/src/arrowTable.js";

import { StepsList } from "./components/StepsList";
import { PreviewGrid } from "./components/PreviewGrid";
import { SchemaView } from "./components/SchemaView";
import { AddStepMenu } from "./components/AddStepMenu";

export type QueryEditorPanelProps = {
  query: Query;
  engine: QueryEngine;
  context?: any;
  onQueryChange?: (next: Query) => void;
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

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setError(null);
        const result = await props.engine.executeQuery(
          { ...props.query, steps: stepsToExecute },
          props.context ?? {},
          { limit: 100 },
        );
        if (!cancelled) setPreview(result);
      } catch (e: any) {
        if (!cancelled) setError(e?.message ?? String(e));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [props.engine, props.query, props.context, stepsToExecute]);

  return (
    <div style={{ display: "grid", gridTemplateColumns: "280px 1fr", height: "100%" }}>
      <div style={{ borderInlineEnd: "1px solid var(--border)", padding: 12, overflow: "auto" }}>
        <h3 style={{ marginTop: 0 }}>{props.query.name}</h3>
        <AddStepMenu
          onAddStep={(operation) => {
            const step: QueryStep = {
              id: crypto.randomUUID(),
              name: operation.type,
              operation,
            };
            const insertAt = effectiveSelectedStepIndex + 1;
            const nextSteps = props.query.steps.slice();
            nextSteps.splice(insertAt, 0, step);
            props.onQueryChange?.({ ...props.query, steps: nextSteps });
            setSelectedStepIndex(insertAt);
          }}
          onAiSuggest={props.onAiSuggestNextSteps}
          aiContext={{ query: props.query, preview }}
        />
        <StepsList steps={props.query.steps} selectedIndex={effectiveSelectedStepIndex} onSelect={setSelectedStepIndex} />
      </div>

      <div style={{ display: "grid", gridTemplateRows: "auto 1fr", overflow: "hidden" }}>
        <div style={{ borderBottom: "1px solid var(--border)", padding: 12 }}>
          <SchemaView table={preview} />
          {error ? <div style={{ color: "var(--error)", marginTop: 8 }}>{error}</div> : null}
        </div>
        <div style={{ overflow: "auto" }}>
          <PreviewGrid table={preview} />
        </div>
      </div>
    </div>
  );
}
