import React from "react";
import type { Table } from "./tableTypes";
import { tWithVars } from "../i18n/index.js";

export interface TableHeaderCellProps {
  table: Table;
  columnIndex: number;
  label: string;
  onOpenFilter?: (columnIndex: number) => void;
}

export function TableHeaderCell({
  table,
  columnIndex,
  label,
  onOpenFilter,
}: TableHeaderCellProps) {
  const style = table.style;

  // A tiny subset of Excel-style table styling: bold header + background.
  // Proper style mapping is out of scope here, but the rendering surface is.
  const hasStyledHeader = Boolean(style?.name.startsWith("TableStyle"));
  const className = hasStyledHeader
    ? "formula-table-header-cell formula-table-header-cell--styled"
    : "formula-table-header-cell";

  const hasFilter = !!table.autoFilter;
  return (
    <div className={className}>
      <span>{label}</span>
      {hasFilter ? (
        <button
          type="button"
          aria-label={tWithVars("table.filter.ariaLabel", { column: label })}
          className="formula-table-filter-button"
          onClick={() => onOpenFilter?.(columnIndex)}
        >
          ▾
        </button>
      ) : null}
    </div>
  );
}
