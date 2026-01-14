import React, { useState } from "react";
import type { SortKey, SortOrder, SortSpec } from "./types";

export type SortDialogProps = {
  columns: { index: number; name: string }[];
  /**
   * Optional fallback column labels used when `hasHeader` is false (Excel-style A/B/C).
   *
   * When omitted, `columns` is used for both modes.
   */
  fallbackColumns?: { index: number; name: string }[];
  initial: SortSpec;
  onCancel: () => void;
  onApply: (spec: SortSpec) => void;
};

function nextOrder(order: SortOrder): SortOrder {
  return order === "ascending" ? "descending" : "ascending";
}

export function SortDialog(props: SortDialogProps) {
  const [keys, setKeys] = useState<SortKey[]>(props.initial.keys);
  const [hasHeader, setHasHeader] = useState<boolean>(props.initial.hasHeader);
  const visibleColumns = hasHeader ? props.columns : props.fallbackColumns ?? props.columns;

  return (
    <div className="formula-sort-dialog" data-testid="sort-dialog">
      <div className="formula-sort-dialog__title">Sort</div>

      <label className="formula-sort-dialog__header-toggle">
        <input
          className="formula-sort-filter__checkbox"
          type="checkbox"
          data-testid="sort-dialog-has-header"
          checked={hasHeader}
          onChange={(e) => setHasHeader(e.target.checked)}
        />{" "}
        My data has headers
      </label>

      <div className="formula-sort-dialog__levels">
        {keys.map((key, i) => (
          <div key={i} className="formula-sort-dialog__level">
            <select
              className="formula-sort-filter__select formula-sort-dialog__select"
              data-testid={`sort-dialog-column-${i}`}
              value={key.column}
              onChange={(e) => {
                const col = Number(e.target.value);
                setKeys((prev) => prev.map((k, idx) => (idx === i ? { ...k, column: col } : k)));
              }}
            >
              {visibleColumns.map((c) => (
                <option key={c.index} value={c.index}>
                  {c.name}
                </option>
              ))}
            </select>
            <button
              className="formula-sort-filter__button"
              data-testid={`sort-dialog-order-${i}`}
              onClick={() => setKeys((prev) => prev.map((k, idx) => (idx === i ? { ...k, order: nextOrder(k.order) } : k)))}
            >
              {key.order === "ascending" ? "A→Z" : "Z→A"}
            </button>
            <button
              className="formula-sort-filter__button"
              data-testid={`sort-dialog-remove-${i}`}
              onClick={() => setKeys((prev) => prev.filter((_, idx) => idx !== i))}
            >
              Remove
            </button>
          </div>
        ))}

        <button
          className="formula-sort-filter__button formula-sort-dialog__add-level"
          data-testid="sort-dialog-add-level"
          onClick={() =>
            setKeys((prev) => [
              ...prev,
              {
                column: visibleColumns[0]?.index ?? 0,
                order: "ascending",
              },
            ])
          }
        >
          Add level
        </button>
      </div>

      <div className="formula-sort-filter__controls formula-sort-dialog__controls">
        <button className="formula-sort-filter__button" data-testid="sort-dialog-cancel" onClick={props.onCancel}>
          Cancel
        </button>
        <button
          className="formula-sort-filter__button formula-sort-filter__button--primary"
          data-testid="sort-dialog-ok"
          onClick={() => props.onApply({ keys, hasHeader })}
          disabled={keys.length === 0}
        >
          OK
        </button>
      </div>
    </div>
  );
}
