import React from "react";

import type { ArrowTableAdapter, DataTable } from "@formula/power-query";

export function SchemaView(props: { table: DataTable | ArrowTableAdapter | null }) {
  if (!props.table) return null;
  return (
    <div className="query-editor-schema">
      {props.table.columns.map((col) => (
        <span key={col.name} className="query-editor-schema__chip">
          {col.name}: {col.type}
        </span>
      ))}
    </div>
  );
}
