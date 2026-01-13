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
    <div className="formula-filter-dropdown">
      <div className="formula-filter-dropdown__title">
        {props.columnName} {props.isFiltered ? t("filterDropdown.filtered") : ""}
      </div>

      <input
        className="formula-sort-filter__input formula-filter-dropdown__search"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder={t("filterDropdown.search.placeholder")}
      />

      <div className="formula-filter-dropdown__values">
        {visibleValues.map((v) => (
          <label key={v} className="formula-sort-filter__row formula-filter-dropdown__value">
            <input className="formula-sort-filter__checkbox" type="checkbox" /> {v}
          </label>
        ))}
      </div>

      <div className="formula-sort-filter__controls formula-filter-dropdown__controls">
        <button className="formula-sort-filter__button" onClick={() => props.onChange(undefined)}>
          {t("filterDropdown.clear")}
        </button>
        <button
          className="formula-sort-filter__button formula-sort-filter__button--primary"
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
