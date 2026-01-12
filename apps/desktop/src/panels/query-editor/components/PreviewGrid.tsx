import React from "react";

import type { ArrowTableAdapter, DataTable } from "@formula/power-query";
import { t } from "../../../i18n/index.js";

export function PreviewGrid(props: { table: DataTable | ArrowTableAdapter | null }) {
  if (!props.table) {
    return <div className="query-editor-preview__empty">{t("queryEditor.preview.none")}</div>;
  }

  const grid = props.table.toGrid({ includeHeader: true });
  return (
    <table className="query-editor-preview__table">
      <thead>
        <tr>
          {grid[0].map((cell, idx) => (
            <th
              key={idx}
              className="query-editor-preview__th"
            >
              {String(cell)}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {grid.slice(1).map((row, rIdx) => (
          <tr key={rIdx}>
            {row.map((cell, cIdx) => (
              <td key={cIdx} className="query-editor-preview__td">
                {cell == null ? "" : String(cell)}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
