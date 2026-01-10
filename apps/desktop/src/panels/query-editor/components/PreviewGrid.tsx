import React from "react";

import type { DataTable } from "../../../../../packages/power-query/src/table.js";

export function PreviewGrid(props: { table: DataTable | null }) {
  if (!props.table) {
    return <div style={{ padding: 12, color: "#666" }}>No preview available.</div>;
  }

  const grid = props.table.toGrid({ includeHeader: true });
  return (
    <table style={{ borderCollapse: "collapse", width: "100%" }}>
      <thead>
        <tr>
          {grid[0].map((cell, idx) => (
            <th key={idx} style={{ position: "sticky", top: 0, background: "#fafafa", borderBottom: "1px solid #ddd", padding: 6 }}>
              {String(cell)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {grid.slice(1).map((row, rIdx) => (
          <tr key={rIdx}>
            {row.map((cell, cIdx) => (
              <td key={cIdx} style={{ borderBottom: "1px solid #eee", padding: 6, fontSize: 12 }}>
                {cell == null ? "" : String(cell)}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

