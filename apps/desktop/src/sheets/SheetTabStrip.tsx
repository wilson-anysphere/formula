import React, { useMemo, useState } from "react";

import type { Sheet, TabColor, Workbook } from "../workbook/workbook";

type Props = {
  workbook: Workbook;
  activeSheetId: string;
  onActivateSheet: (sheetId: string) => void;
};

export function SheetTabStrip({ workbook, activeSheetId, onActivateSheet }: Props) {
  const visibleSheets = useMemo(
    () => workbook.sheets.filter((s) => s.visibility === "visible"),
    [workbook.sheets],
  );

  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState("");

  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        borderTop: "1px solid var(--border)",
        padding: "4px 8px",
        gap: 6,
        userSelect: "none",
      }}
      onKeyDown={(e) => {
        if (!e.ctrlKey) return;
        if (e.key !== "PageUp" && e.key !== "PageDown") return;
        e.preventDefault();

        const idx = visibleSheets.findIndex((s) => s.id === activeSheetId);
        if (idx === -1) return;
        const delta = e.key === "PageUp" ? -1 : 1;
        const next = visibleSheets[(idx + delta + visibleSheets.length) % visibleSheets.length];
        if (next) onActivateSheet(next.id);
      }}
      tabIndex={0}
    >
      {visibleSheets.map((sheet) => (
        <SheetTab
          key={sheet.id}
          sheet={sheet}
          active={sheet.id === activeSheetId}
          editing={editingSheetId === sheet.id}
          draftName={draftName}
          onActivate={() => onActivateSheet(sheet.id)}
          onBeginRename={() => {
            setEditingSheetId(sheet.id);
            setDraftName(sheet.name);
          }}
          onCommitRename={() => {
            workbook.renameSheet(sheet.id, draftName);
            setEditingSheetId(null);
          }}
          onCancelRename={() => setEditingSheetId(null)}
          onDraftNameChange={setDraftName}
          onReorder={(targetId) => {
            const from = workbook.sheets.findIndex((s) => s.id === sheet.id);
            const to = workbook.sheets.findIndex((s) => s.id === targetId);
            if (from !== -1 && to !== -1) workbook.reorderSheet(sheet.id, to);
          }}
        />
      ))}

      <button
        type="button"
        data-testid="sheet-add"
        onClick={() => {
          const sheet = workbook.addSheet();
          onActivateSheet(sheet.id);
        }}
        style={{
          height: 24,
          width: 28,
          border: "1px solid var(--border)",
          borderRadius: 4,
          background: "var(--bg-primary)",
          cursor: "pointer",
        }}
        aria-label="Add sheet"
      >
        +
      </button>
    </div>
  );
}

function SheetTab(props: {
  sheet: Sheet;
  active: boolean;
  editing: boolean;
  draftName: string;
  onActivate: () => void;
  onBeginRename: () => void;
  onCommitRename: () => void;
  onCancelRename: () => void;
  onDraftNameChange: (name: string) => void;
  onReorder: (targetId: string) => void;
}) {
  const { sheet, active, editing, draftName } = props;

  return (
    <div
      draggable={!editing}
      onDragStart={(e) => {
        e.dataTransfer.setData("text/sheet-id", sheet.id);
        e.dataTransfer.effectAllowed = "move";
      }}
      onDragOver={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id")) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
      }}
      onDrop={(e) => {
        const fromId = e.dataTransfer.getData("text/sheet-id");
        if (fromId && fromId !== sheet.id) props.onReorder(sheet.id);
      }}
      onClick={() => props.onActivate()}
      onDoubleClick={() => props.onBeginRename()}
      data-testid={`sheet-tab-${sheet.name}`}
      style={{
        padding: "4px 10px",
        borderRadius: 6,
        border: `1px solid ${active ? "var(--accent)" : "var(--border)"}`,
        background: active ? "var(--accent-light)" : "var(--bg-primary)",
        cursor: "pointer",
        position: "relative",
      }}
    >
      {editing ? (
        <input
          value={draftName}
          autoFocus
          onChange={(e) => props.onDraftNameChange(e.target.value)}
          onBlur={() => props.onCommitRename()}
          onKeyDown={(e) => {
            if (e.key === "Enter") props.onCommitRename();
            if (e.key === "Escape") props.onCancelRename();
          }}
          style={{
            width: Math.max(60, draftName.length * 8),
            font: "inherit",
            border: "1px solid var(--accent)",
          }}
        />
      ) : (
        <>
          <span>{sheet.name}</span>
          {sheet.tabColor?.rgb ? <TabColorUnderline tabColor={sheet.tabColor} /> : null}
        </>
      )}
    </div>
  );
}

function TabColorUnderline({ tabColor }: { tabColor: TabColor }) {
  const rgb = tabColor.rgb;
  if (!rgb) return null;
  const css = argbToCssHsl(rgb);
  return (
    <div
      style={{
        position: "absolute",
        left: 6,
        right: 6,
        bottom: 2,
        height: 3,
        borderRadius: 2,
        background: css,
      }}
    />
  );
}

function argbToCssHsl(argb: string): string {
  if (!/^([0-9A-Fa-f]{8})$/.test(argb)) return "transparent";
  const alpha = parseInt(argb.slice(0, 2), 16) / 255;
  const r = parseInt(argb.slice(2, 4), 16);
  const g = parseInt(argb.slice(4, 6), 16);
  const b = parseInt(argb.slice(6, 8), 16);

  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;
  const light = (max + min) / 2;

  let hue = 0;
  let sat = 0;

  if (delta !== 0) {
    sat = delta / (1 - Math.abs(2 * light - 1));
    switch (max) {
      case rn:
        hue = ((gn - bn) / delta + (gn < bn ? 6 : 0)) * 60;
        break;
      case gn:
        hue = ((bn - rn) / delta + 2) * 60;
        break;
      default:
        hue = ((rn - gn) / delta + 4) * 60;
        break;
    }
  }

  const h = Math.round(hue);
  const s = Math.round(sat * 100);
  const l = Math.round(light * 100);
  const a = Math.max(0, Math.min(1, alpha));

  return a < 1 ? `hsla(${h}, ${s}%, ${l}%, ${a.toFixed(3)})` : `hsl(${h}, ${s}%, ${l}%)`;
}
