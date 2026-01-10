import React from "react";
import type { Table } from "./tableTypes";

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
  const background =
    style?.name.startsWith("TableStyle") ? "var(--bg-secondary)" : "transparent";

  const hasFilter = !!table.autoFilter;
  return (
    <div
      className="formula-table-header-cell"
      style={{
        fontWeight: 600,
        background,
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 6,
        padding: "0 6px",
        height: "100%",
        boxSizing: "border-box",
      }}
    >
      <span>{label}</span>
      {hasFilter ? (
        <button
          type="button"
          aria-label={`Filter ${label}`}
          className="formula-table-filter-button"
          onClick={() => onOpenFilter?.(columnIndex)}
        >
          ▾
        </button>
      ) : null}
    </div>
  );
}
