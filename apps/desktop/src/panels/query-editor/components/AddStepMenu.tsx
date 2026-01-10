import React, { useState } from "react";

import type { Query, QueryOperation } from "../../../../../packages/power-query/src/model.js";
import type { DataTable } from "../../../../../packages/power-query/src/table.js";

export function AddStepMenu(props: {
  onAddStep: (op: QueryOperation) => void;
  onAiSuggest?: (intent: string, ctx: { query: Query; preview: DataTable | null }) => Promise<QueryOperation[]>;
  aiContext: { query: Query; preview: DataTable | null };
}) {
  const [intent, setIntent] = useState("");
  const [suggestions, setSuggestions] = useState<QueryOperation[] | null>(null);

  return (
    <div style={{ marginBottom: 12 }}>
      <button
        type="button"
        onClick={() => props.onAddStep({ type: "filterRows", predicate: { type: "comparison", column: "", operator: "isNotNull" } } as any)}
        style={{ width: "100%", marginBottom: 8 }}
      >
        + Add step (starter)
      </button>

      {props.onAiSuggest ? (
        <div>
          <input
            value={intent}
            onChange={(e) => setIntent(e.target.value)}
            placeholder="Ask AI: e.g. 'filter to East, group by Region'"
            style={{ width: "100%", boxSizing: "border-box", marginBottom: 6 }}
          />
          <button
            type="button"
            onClick={async () => {
              const ops = await props.onAiSuggest?.(intent, props.aiContext);
              setSuggestions(ops ?? []);
            }}
            disabled={!intent.trim()}
            style={{ width: "100%" }}
          >
            Suggest next steps
          </button>
          {suggestions ? (
            <div style={{ marginTop: 8 }}>
              {suggestions.length === 0 ? (
                <div style={{ fontSize: 12, color: "#666" }}>No suggestions.</div>
              ) : (
                suggestions.map((op, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => props.onAddStep(op)}
                    style={{ width: "100%", marginTop: 4, textAlign: "left" }}
                  >
                    {op.type}
                  </button>
                ))
              )}
            </div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

