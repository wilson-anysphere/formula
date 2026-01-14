import React, { useEffect, useMemo, useRef, useState } from "react";
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
  const allSelected = values.length > 0 && selected.size === values.length;
  const noneSelected = selected.size === 0;
  const selectAllRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    const el = selectAllRef.current;
    if (!el) return;
    el.indeterminate = !allSelected && !noneSelected;
  }, [allSelected, noneSelected]);

  const toggle = (v: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(v)) next.delete(v);
      else next.add(v);
      return next;
    });
  };

  const toggleAll = () => {
    setSelected(() => (allSelected ? new Set() : new Set(values)));
  };

  return (
    <div className="formula-table-filter-dropdown">
      <div className="formula-table-filter-dropdown__list">
        {values.length > 0 && (
          <label key="__select_all__" className="formula-sort-filter__row formula-table-filter-dropdown__item">
            <input
              ref={selectAllRef}
              className="formula-sort-filter__checkbox"
              type="checkbox"
              checked={allSelected}
              onChange={toggleAll}
            />
            <span>Select All</span>
          </label>
        )}
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
