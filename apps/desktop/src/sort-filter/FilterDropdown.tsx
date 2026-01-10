import React, { useMemo, useState } from "react";
import type { ColumnFilter } from "./types";
import { t } from "../i18n/index.js";

export type FilterDropdownProps = {
  columnName: string;
  uniqueValues: string[];
  value: ColumnFilter | undefined;
  onChange: (next: ColumnFilter | undefined) => void;
  isFiltered: boolean;
};

export function FilterDropdown(props: FilterDropdownProps) {
  const [query, setQuery] = useState("");
  const visibleValues = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return props.uniqueValues;
    return props.uniqueValues.filter((v) => v.toLowerCase().includes(q));
  }, [props.uniqueValues, query]);

  return (
    <div style={{ width: 280, padding: 8 }}>
      <div style={{ fontWeight: 600, marginBottom: 8 }}>
        {props.columnName} {props.isFiltered ? t("filterDropdown.filtered") : ""}
      </div>

      <input
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder={t("filterDropdown.search.placeholder")}
        style={{ width: "100%", marginBottom: 8 }}
      />

      <div style={{ maxHeight: 220, overflow: "auto", border: "1px solid var(--border)" }}>
        {visibleValues.map((v) => (
          <label key={v} style={{ display: "block", padding: "4px 8px" }}>
            <input type="checkbox" /> {v}
          </label>
        ))}
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 8, justifyContent: "flex-end" }}>
        <button onClick={() => props.onChange(undefined)}>{t("filterDropdown.clear")}</button>
        <button
          onClick={() =>
            props.onChange({
              join: "any",
              criteria: [],
            })
          }
        >
          {t("filterDropdown.apply")}
        </button>
      </div>
    </div>
  );
}
