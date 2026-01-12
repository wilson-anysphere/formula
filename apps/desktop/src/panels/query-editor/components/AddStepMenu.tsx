import React, { useState } from "react";

import type { ArrowTableAdapter, DataTable, Query, QueryOperation } from "@formula/power-query";
import { t } from "../../../i18n/index.js";

export function AddStepMenu(props: {
  onAddStep: (op: QueryOperation) => void;
  onAiSuggest?: (intent: string, ctx: { query: Query; preview: DataTable | ArrowTableAdapter | null }) => Promise<QueryOperation[]>;
  aiContext: { query: Query; preview: DataTable | ArrowTableAdapter | null };
}) {
  const [intent, setIntent] = useState("");
  const [suggestions, setSuggestions] = useState<QueryOperation[] | null>(null);

  return (
    <div className="query-editor-add-step">
      <button
        type="button"
        onClick={() => props.onAddStep({ type: "filterRows", predicate: { type: "comparison", column: "", operator: "isNotNull" } } as any)}
        className="query-editor-add-step__starter"
      >
        {t("queryEditor.addStep.addStarter")}
      </button>

      {props.onAiSuggest ? (
        <div>
          <input
            value={intent}
            onChange={(e) => setIntent(e.target.value)}
            placeholder={t("queryEditor.addStep.aiPlaceholder")}
            className="query-editor-add-step__ai-input"
          />
          <button
            type="button"
            onClick={async () => {
              const ops = await props.onAiSuggest?.(intent, props.aiContext);
              setSuggestions(ops ?? []);
            }}
            disabled={!intent.trim()}
            className="query-editor-add-step__ai-button"
          >
            {t("queryEditor.addStep.suggestNext")}
          </button>
          {suggestions ? (
            <div className="query-editor-add-step__suggestions">
              {suggestions.length === 0 ? (
                <div className="query-editor-add-step__no-suggestions">{t("queryEditor.addStep.noSuggestions")}</div>
              ) : (
                suggestions.map((op, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => props.onAddStep(op)}
                    className="query-editor-add-step__suggestion"
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
