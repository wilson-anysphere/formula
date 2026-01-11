import type { GridApi } from "@formula/grid";
import { useLayoutEffect, useRef, useState } from "react";

type ViewportRect = { x: number; y: number; width: number; height: number };

function isSameRect(a: ViewportRect | null, b: ViewportRect | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
}

export function CellEditorOverlay(props: {
  gridApi: GridApi | null;
  cell: { row: number; col: number } | null;
  value: string;
  onChange: (value: string) => void;
  onCommit: (nav: { deltaRow: number; deltaCol: number }) => void;
  onCancel: () => void;
}): JSX.Element | null {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const focusedCellKeyRef = useRef<string | null>(null);
  const [rect, setRect] = useState<ViewportRect | null>(null);

  useLayoutEffect(() => {
    if (!props.gridApi || !props.cell) return;

    let frame = 0;
    const tick = () => {
      const next = props.gridApi?.getCellRect(props.cell!.row, props.cell!.col) ?? null;
      setRect((prev) => (isSameRect(prev, next) ? prev : next));
      frame = requestAnimationFrame(tick);
    };

    tick();
    return () => cancelAnimationFrame(frame);
  }, [props.gridApi, props.cell?.row, props.cell?.col]);

  useLayoutEffect(() => {
    if (!props.cell || !rect) return;
    const input = inputRef.current;
    if (!input) return;
    const key = `${props.cell.row}:${props.cell.col}`;
    if (focusedCellKeyRef.current === key) return;
    focusedCellKeyRef.current = key;
    input.focus({ preventScroll: true });
    const end = input.value.length;
    input.setSelectionRange(end, end);
  }, [props.cell, rect]);

  if (!props.cell || !rect) return null;

  return (
    <input
      ref={inputRef}
      data-testid="cell-editor"
      spellCheck={false}
      value={props.value}
      onChange={(event) => props.onChange(event.currentTarget.value)}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          props.onCancel();
          return;
        }

        if (event.key === "Enter") {
          event.preventDefault();
          props.onCommit({ deltaRow: event.shiftKey ? -1 : 1, deltaCol: 0 });
          return;
        }

        if (event.key === "Tab") {
          event.preventDefault();
          props.onCommit({ deltaRow: 0, deltaCol: event.shiftKey ? -1 : 1 });
        }
      }}
      style={{
        position: "absolute",
        left: rect.x,
        top: rect.y,
        width: rect.width,
        height: rect.height,
        boxSizing: "border-box",
        zIndex: 10,
        border: "2px solid var(--formula-grid-selection-border, #0e65eb)",
        borderRadius: 2,
        padding: "0 6px",
        fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
        fontSize: 13,
        background: "var(--formula-grid-bg, #ffffff)",
        color: "var(--formula-grid-cell-text, #0f172a)"
      }}
    />
  );
}
