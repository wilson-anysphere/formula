import React, { useMemo, useState } from "react";
import { distinctColumnValues, TableViewRow } from "./tableView";

export interface AutoFilterDropdownProps {
  rows: TableViewRow[];
  colId: number;
  /**
   * Initial selected values.
   *
   * - `null`/`undefined` means "select all" (default Excel-like behavior when no filter exists yet).
   * - An explicit empty array means "select none" (show nothing).
   */
  initialSelected?: string[] | null;
  onApply: (selected: string[]) => void;
  onClose: () => void;
}

export function AutoFilterDropdown({
  rows,
  colId,
  initialSelected,
  onApply,
  onClose,
}: AutoFilterDropdownProps) {
  const values = useMemo(() => distinctColumnValues(rows, colId), [rows, colId]);
  const [selected, setSelected] = useState<Set<string>>(
    () => new Set(initialSelected == null ? values : initialSelected),
  );

  const valueLabel = (value: string): string => (value === "" ? "(Blanks)" : value);

  const toggle = (v: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(v)) next.delete(v);
      else next.add(v);
      return next;
    });
  };

  return (
    <div className="formula-table-filter-dropdown">
      <div className="formula-table-filter-dropdown__list">
        {values.map((v) => (
          <label key={v} className="formula-sort-filter__row formula-table-filter-dropdown__item">
            <input
              className="formula-sort-filter__checkbox"
              type="checkbox"
              checked={selected.has(v)}
              onChange={() => toggle(v)}
            />
            <span>{valueLabel(v)}</span>
          </label>
        ))}
      </div>
      <div className="formula-sort-filter__controls formula-table-filter-dropdown__controls">
        <button className="formula-sort-filter__button" type="button" onClick={onClose}>
          Cancel
        </button>
        <button
          className="formula-sort-filter__button formula-sort-filter__button--primary"
          type="button"
          onClick={() => {
            onApply(Array.from(selected));
            onClose();
          }}
        >
          Apply
        </button>
      </div>
    </div>
  );
}
