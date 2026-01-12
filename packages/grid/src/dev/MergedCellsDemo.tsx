import React, { useEffect, useMemo, useRef, useState } from "react";

import type { CellData, CellProvider, CellRange, CellStyle } from "../model/CellProvider";
import { CanvasGrid, type GridApi } from "../react/CanvasGrid";

function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;
}

function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
}

class MergedDemoProvider implements CellProvider {
  private readonly mergedRanges: CellRange[];

  constructor() {
    this.mergedRanges = [
      { startRow: 1, endRow: 3, startCol: 1, endCol: 4 },
      { startRow: 5, endRow: 6, startCol: 2, endCol: 6 }
    ];
  }

  getMergedRangeAt(row: number, col: number): CellRange | null {
    for (const range of this.mergedRanges) {
      if (rangeContains(range, row, col)) return range;
    }
    return null;
  }

  getMergedRangesInRange(range: CellRange): CellRange[] {
    return this.mergedRanges.filter((merged) => rangesIntersect(merged, range));
  }

  getCell(row: number, col: number): CellData | null {
    // Avoid hard-coded header colors so the demo respects `GridTheme.headerBg/headerText`
    // (and the CSS variable theme examples in `packages/grid/index.html`).
    const headerStyle: CellStyle = { fontWeight: "600" };
    if (row === 0) return { row, col, value: col === 0 ? "" : `Col ${col}`, style: headerStyle };
    if (col === 0) return { row, col, value: row, style: headerStyle };

    // Merged cell examples (anchor at top-left).
    if (row === 1 && col === 1) {
      return {
        row,
        col,
        value: "Merged header (B2:D3)",
        style: { fill: "#dbeafe", fontWeight: "600", textAlign: "center" }
      };
    }

    if (row === 5 && col === 2) {
      return {
        row,
        col,
        value: "Merged + overflow →",
        style: { fill: "#fef3c7", fontWeight: "600", textAlign: "start", wrapMode: "none" }
      };
    }

    // Text overflow examples.
    if (row === 4 && col === 2) {
      return {
        row,
        col,
        value: "Left-aligned text overflows into empty neighbors until it hits a non-empty cell.",
        style: { wrapMode: "none", textAlign: "start" }
      };
    }
    if (row === 4 && col === 7) {
      return { row, col, value: "STOP", style: { fill: "#fee2e2", fontWeight: "600" } };
    }

    if (row === 7 && col === 7) {
      return {
        row,
        col,
        value: "Right-aligned overflow ←",
        style: { wrapMode: "none", textAlign: "end", fontWeight: "600" }
      };
    }
    if (row === 7 && col === 4) {
      return { row, col, value: "STOP", style: { fill: "#dcfce7", fontWeight: "600" } };
    }

    return null;
  }
}

export function MergedCellsDemo(): React.ReactElement {
  const provider = useMemo(() => new MergedDemoProvider(), []);
  const apiRef = useRef<GridApi | null>(null);
  const [zoom, setZoom] = useState(1);

  useEffect(() => {
    apiRef.current?.setZoom(zoom);
  }, [zoom]);

  return (
    <div style={{ width: "100%", height: "100%", display: "flex", flexDirection: "column" }}>
      <div
        style={{
          padding: 12,
          borderBottom: "1px solid var(--formula-grid-line, #e5e7eb)",
          background: "var(--formula-grid-header-bg, #fff)",
          color: "var(--formula-grid-header-text, #0f172a)",
          fontFamily: "system-ui, sans-serif",
          fontSize: 14
        }}
      >
        <div style={{ fontWeight: 600 }}>Merged cells + Excel-style text overflow</div>
        <div style={{ marginTop: 4, color: "var(--formula-grid-cell-text, #4b5563)", opacity: 0.75 }}>
          Try clicking inside merged regions (selection snaps to the anchor) and observe text overflowing into empty neighbors.
          Append <code>?demo=perf</code> to the URL to switch back to the perf harness.
        </div>
        <label style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10 }}>
          Zoom
          <input
            type="range"
            min={0.25}
            max={4}
            step={0.1}
            value={zoom}
            onChange={(event) => setZoom(event.currentTarget.valueAsNumber)}
          />
          <span style={{ width: 44, textAlign: "right" }}>{Math.round(zoom * 100)}%</span>
        </label>
      </div>

      <div style={{ flex: 1, position: "relative" }}>
        <CanvasGrid
          provider={provider}
          rowCount={40}
          colCount={15}
          headerRows={1}
          headerCols={1}
          frozenRows={1}
          frozenCols={1}
          defaultRowHeight={24}
          defaultColWidth={80}
          onZoomChange={setZoom}
          apiRef={apiRef}
        />
      </div>
    </div>
  );
}
