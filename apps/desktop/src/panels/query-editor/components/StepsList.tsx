import React from "react";

import type { QueryStep } from "@formula/power-query";
import { t } from "../../../i18n/index.js";

export function StepsList(props: { steps: QueryStep[]; selectedIndex: number; onSelect: (idx: number) => void }) {
  return (
    <div>
      <div className="query-editor-steps__title">{t("queryEditor.steps.title")}</div>
      <ol className="query-editor-steps__list">
        {props.steps.map((step, idx) => (
          <li key={step.id}>
            <button
              type="button"
              onClick={() => props.onSelect(idx)}
              className={
                idx === props.selectedIndex
                  ? "query-editor-steps__button query-editor-steps__button--selected"
                  : "query-editor-steps__button"
              }
            >
              {step.name}
            </button>
          </li>
        ))}
      </ol>
    </div>
  );
}
