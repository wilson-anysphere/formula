import React, { useState } from "react";
import type { SortKey, SortOrder, SortSpec } from "./types";

export type SortDialogProps = {
  columns: { index: number; name: string }[];
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

  return (
    <div style={{ width: 420, padding: 12 }}>
      <div style={{ fontWeight: 600, marginBottom: 8 }}>Sort</div>

      <label style={{ display: "block", marginBottom: 12 }}>
        <input type="checkbox" checked={hasHeader} onChange={(e) => setHasHeader(e.target.checked)} />{" "}
        My data has headers
      </label>

      <div style={{ display: "grid", gap: 8 }}>
        {keys.map((key, i) => (
          <div key={i} style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <select
              value={key.column}
              onChange={(e) => {
                const col = Number(e.target.value);
                setKeys((prev) => prev.map((k, idx) => (idx === i ? { ...k, column: col } : k)));
              }}
            >
              {props.columns.map((c) => (
                <option key={c.index} value={c.index}>
                  {c.name}
                </option>
              ))}
            </select>
            <button
              onClick={() => setKeys((prev) => prev.map((k, idx) => (idx === i ? { ...k, order: nextOrder(k.order) } : k)))}
            >
              {key.order === "ascending" ? "A→Z" : "Z→A"}
            </button>
            <button onClick={() => setKeys((prev) => prev.filter((_, idx) => idx !== i))}>Remove</button>
          </div>
        ))}

        <button
          onClick={() =>
            setKeys((prev) => [
              ...prev,
              {
                column: props.columns[0]?.index ?? 0,
                order: "ascending",
              },
            ])
          }
        >
          Add level
        </button>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button onClick={props.onCancel}>Cancel</button>
        <button onClick={() => props.onApply({ keys, hasHeader })} disabled={keys.length === 0}>
          OK
        </button>
      </div>
    </div>
  );
}

