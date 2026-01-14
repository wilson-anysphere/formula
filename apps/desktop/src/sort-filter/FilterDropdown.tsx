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
  const blanksLabel = t("filterDropdown.blanks");
  const valueLabel = (value: string): string => (value === "" ? blanksLabel : value);
  const visibleValues = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return props.uniqueValues;
    const blank = blanksLabel.toLowerCase();
    return props.uniqueValues.filter((v) => (v === "" ? blank : v).toLowerCase().includes(q));
  }, [props.uniqueValues, query, blanksLabel]);

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
        aria-label={t("filterDropdown.search.ariaLabel")}
      />

      <div className="formula-filter-dropdown__values">
        {visibleValues.length === 0 ? (
          <div className="formula-table-filter-dropdown__empty">{t("filterDropdown.noMatches")}</div>
        ) : null}
        {visibleValues.map((v) => (
          <label key={v} className="formula-sort-filter__row formula-filter-dropdown__value">
            <input className="formula-sort-filter__checkbox" type="checkbox" /> {valueLabel(v)}
          </label>
        ))}
      </div>

      <div className="formula-sort-filter__controls formula-filter-dropdown__controls">
        <button className="formula-sort-filter__button" type="button" onClick={() => props.onChange(undefined)}>
          {t("filterDropdown.clear")}
        </button>
        <button
          className="formula-sort-filter__button formula-sort-filter__button--primary"
          type="button"
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
