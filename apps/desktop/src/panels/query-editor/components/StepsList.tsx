import React from "react";

import type { QueryStep } from "@formula/power-query";
import { t } from "../../../i18n/index.js";

export function StepsList(props: { steps: QueryStep[]; selectedIndex: number; onSelect: (idx: number) => void }) {
  return (
    <div>
      <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}>{t("queryEditor.steps.title")}</div>
      <ol style={{ listStyle: "none", padding: 0, margin: 0 }}>
        {props.steps.map((step, idx) => (
          <li key={step.id}>
            <button
              type="button"
              onClick={() => props.onSelect(idx)}
              style={{
                width: "100%",
                textAlign: "start",
                border: "none",
                background: idx === props.selectedIndex ? "var(--selection-bg)" : "transparent",
                padding: "6px 8px",
                cursor: "pointer",
                borderRadius: 4,
              }}
            >
              {step.name}
            </button>
          </li>
        ))}
      </ol>
    </div>
  );
}
