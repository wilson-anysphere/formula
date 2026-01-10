import React, { useMemo, useState } from "react";
import { distinctColumnValues, TableViewRow } from "./tableView";

export interface AutoFilterDropdownProps {
  rows: TableViewRow[];
  colId: number;
  initialSelected: string[];
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
    () => new Set(initialSelected.length ? initialSelected : values),
  );

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
      <div style={{ maxHeight: 240, overflow: "auto", padding: 6 }}>
        {values.map((v) => (
          <label key={v} style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <input
              type="checkbox"
              checked={selected.has(v)}
              onChange={() => toggle(v)}
            />
            <span>{v}</span>
          </label>
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, padding: 6 }}>
        <button type="button" onClick={onClose}>
          Cancel
        </button>
        <button
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

