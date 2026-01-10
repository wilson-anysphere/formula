import React from "react";

import type { DataTable } from "../../../../../packages/power-query/src/table.js";

export function SchemaView(props: { table: DataTable | null }) {
  if (!props.table) return null;
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
      {props.table.columns.map((col) => (
        <span key={col.name} style={{ fontSize: 12, background: "#f2f2f2", padding: "2px 6px", borderRadius: 4 }}>
          {col.name}: {col.type}
        </span>
      ))}
    </div>
  );
}

