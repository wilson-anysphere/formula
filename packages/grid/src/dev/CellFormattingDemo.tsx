import React, { useEffect, useMemo, useRef, useState } from "react";

import type {
  CellBorderLineStyle,
  CellBorderSpec,
  CellBorders,
  CellData,
  CellProvider,
  CellStyle
} from "../model/CellProvider";
import { CanvasGrid, type GridApi } from "../react/CanvasGrid";

function toColumnName(col0: number): string {
  let value = col0 + 1;
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

class CellFormattingDemoProvider implements CellProvider {
  private readonly rowCount: number;
  private readonly colCount: number;
  private readonly cells = new Map<string, CellData>();

  private readonly headerStyle: CellStyle = { fontWeight: "600", textAlign: "center" };
  private readonly rowHeaderStyle: CellStyle = { fontWeight: "600", textAlign: "end" };

  constructor(options: { rowCount: number; colCount: number }) {
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;

    const put = (
      row: number,
      col: number,
      value: CellData["value"],
      style?: CellStyle,
      extras?: Partial<CellData>
    ) => {
      const cell: CellData = { ...(extras ?? {}), row, col, value, style };
      this.cells.set(cellKey(row, col), cell);
    };

    const sectionHeaderFill = "rgba(14, 101, 235, 0.10)";

    // ---------------------------------------------------------------------------------------------
    // Font styles (bold / italic / underline / strike) — these use extra fields so the demo can
    // validate new formatting features as they land (Excel-style rendering).
    // ---------------------------------------------------------------------------------------------
    put(1, 1, "Text styles", { fill: sectionHeaderFill, fontWeight: "700" });

    const baseTextStyle: CellStyle = {
      fill: "rgba(148, 163, 184, 0.08)",
      textAlign: "center"
    };

    put(2, 1, "Bold", {
      ...baseTextStyle,
      fontWeight: "700"
    });

    put(2, 2, "Italic", {
      ...baseTextStyle,
      fontStyle: "italic"
    });

    put(2, 3, "Underline", {
      ...baseTextStyle,
      underline: true
    });

    put(2, 4, "Strike", {
      ...baseTextStyle,
      strike: true
    });

    put(3, 1, "Bold+Italic", {
      ...baseTextStyle,
      fontWeight: "700",
      fontStyle: "italic"
    });

    put(3, 2, "Bold+Underline", {
      ...baseTextStyle,
      fontWeight: "700",
      underline: true
    });

    put(3, 3, "Italic+Underline", {
      ...baseTextStyle,
      fontStyle: "italic",
      underline: true
    });

    put(3, 4, "All", {
      ...baseTextStyle,
      fontWeight: "700",
      fontStyle: "italic",
      underline: true,
      strike: true
    });

    put(4, 1, "Underline+Strike", {
      ...baseTextStyle,
      underline: true,
      strike: true
    });

    put(4, 2, "Italic+Strike", {
      ...baseTextStyle,
      fontStyle: "italic",
      strike: true
    });

    put(4, 3, "Bold+Strike", {
      ...baseTextStyle,
      fontWeight: "700",
      strike: true
    });

    put(4, 4, "Underline+Italic", {
      ...baseTextStyle,
      fontStyle: "italic",
      underline: true
    });

    // ---------------------------------------------------------------------------------------------
    // Fills + font color
    // ---------------------------------------------------------------------------------------------
    put(5, 1, "Fill + font color", { fill: sectionHeaderFill, fontWeight: "700" });

    put(6, 1, "Light fill", { fill: "#dbeafe", color: "#1e3a8a", fontWeight: "600" });
    put(6, 2, "Dark fill", { fill: "#1e3a8a", color: "#ffffff", fontWeight: "600" });
    put(6, 3, "Custom font", {
      fill: "#ecfccb",
      color: "#14532d",
      fontWeight: "600",
      fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, \"Liberation Mono\", \"Courier New\", monospace"
    });

    // ---------------------------------------------------------------------------------------------
    // Borders — intentionally uses "Excel-like" border objects (thin/medium/thick, dashed/dotted/double),
    // including conflicts across shared edges.
    // ---------------------------------------------------------------------------------------------
    put(8, 1, "Borders", { fill: sectionHeaderFill, fontWeight: "700" });

    const borderCellBase: CellStyle = { textAlign: "center", fill: "rgba(148, 163, 184, 0.06)" };
    const mkBorder = (width: number, style: CellBorderLineStyle, color: string): CellBorderSpec => ({ width, style, color });
    const allBorders = (spec: CellBorderSpec): CellBorders => ({ top: spec, right: spec, bottom: spec, left: spec });

    put(9, 1, "Thin", {
      ...borderCellBase,
      borders: allBorders(mkBorder(1, "solid", "#0f172a"))
    });

    put(9, 2, "Medium", {
      ...borderCellBase,
      borders: allBorders(mkBorder(2, "solid", "#0f172a"))
    });

    put(9, 3, "Thick", {
      ...borderCellBase,
      borders: allBorders(mkBorder(3, "solid", "#0f172a"))
    });

    put(10, 1, "Dashed", {
      ...borderCellBase,
      borders: allBorders(mkBorder(1, "dashed", "#ef4444"))
    });

    put(10, 2, "Dotted", {
      ...borderCellBase,
      borders: allBorders(mkBorder(1, "dotted", "#3b82f6"))
    });

    put(10, 3, "Double", {
      ...borderCellBase,
      borders: allBorders(mkBorder(2, "double", "#a855f7"))
    });

    // Conflicting adjacent borders (shared edge). The renderer should pick the correct winner.
    put(11, 1, "Conflict A", {
      ...borderCellBase,
      borders: { right: mkBorder(3, "solid", "#ef4444") }
    });
    put(11, 2, "Conflict B", {
      ...borderCellBase,
      borders: { left: mkBorder(1, "dotted", "#3b82f6") }
    });

    // Conflicting top/bottom border (stacked).
    put(12, 3, "Bottom thick", {
      ...borderCellBase,
      borders: { bottom: mkBorder(3, "solid", "#22c55e") }
    });
    put(13, 3, "Top thin", {
      ...borderCellBase,
      borders: { top: mkBorder(1, "solid", "#f97316") }
    });

    // ---------------------------------------------------------------------------------------------
    // Alignment, wrapping, rotation
    // ---------------------------------------------------------------------------------------------
    put(14, 1, "Alignment / wrap / rotation", { fill: sectionHeaderFill, fontWeight: "700" });

    put(15, 1, "Left", { textAlign: "start" });
    put(15, 2, "Center", { textAlign: "center" });
    put(15, 3, "Right", { textAlign: "end" });

    put(16, 1, "Wrap off: This text should overflow →", { wrapMode: "none", textAlign: "start" });
    put(16, 2, "Wrap on: This text should wrap within the cell (word wrap).", { wrapMode: "word", textAlign: "start" });

    put(17, 1, "Rotated 45°", {
      rotationDeg: 45,
      textAlign: "center",
      verticalAlign: "middle",
      fill: "rgba(14, 101, 235, 0.12)",
      fontWeight: "600"
    });

    // ---------------------------------------------------------------------------------------------
    // Rich text runs (if Task 84 lands)
    // ---------------------------------------------------------------------------------------------
    const richTextValue = "Rich: bold + italic + underline + color";
    put(
      20,
      1,
      richTextValue,
      {
        wrapMode: "word",
        textAlign: "start",
        fill: "rgba(148, 163, 184, 0.08)"
      },
      {
        richText: {
          text: richTextValue,
          runs: [
            { start: 0, end: 6, style: {} },
            { start: 6, end: 10, style: { bold: true } },
            { start: 10, end: 13, style: {} },
            { start: 13, end: 19, style: { italic: true } },
            { start: 19, end: 22, style: {} },
            { start: 22, end: 31, style: { underline: "double", color: "#FF2563EB" } },
            { start: 31, end: 34, style: {} },
            { start: 34, end: 39, style: { color: "#FFEF4444", bold: true } }
          ]
        }
      }
    );

    // ---------------------------------------------------------------------------------------------
    // Number format display strings (Task 68)
    // ---------------------------------------------------------------------------------------------
    put(22, 1, "Number formats", { fill: sectionHeaderFill, fontWeight: "700" });
    const formatCellStyle: CellStyle = { textAlign: "end", fontWeight: "600", fill: "rgba(148, 163, 184, 0.08)" };
    put(23, 1, "$1,234.56", formatCellStyle);
    put(23, 2, "50%", formatCellStyle);
    put(23, 3, "1/2/2024", formatCellStyle);
  }

  prefetch(): void {
    // Synchronous demo provider.
  }

  getCell(row: number, col: number): CellData | null {
    if (row < 0 || col < 0 || row >= this.rowCount || col >= this.colCount) return null;

    // Header cells: avoid hard-coded colors so the demo respects `GridTheme.headerBg/headerText`.
    if (row === 0) {
      return {
        row,
        col,
        value: col === 0 ? "" : toColumnName(col - 1),
        style: this.headerStyle
      };
    }

    if (col === 0) {
      return {
        row,
        col,
        value: row,
        style: this.rowHeaderStyle
      };
    }

    return this.cells.get(cellKey(row, col)) ?? null;
  }
}

export function CellFormattingDemo(): React.ReactElement {
  const rowCount = 30;
  const colCount = 15;

  const provider = useMemo(() => new CellFormattingDemoProvider({ rowCount, colCount }), [rowCount, colCount]);
  const apiRef = useRef<GridApi | null>(null);
  const [zoom, setZoom] = useState(1);

  useEffect(() => {
    apiRef.current?.setZoom(zoom);
  }, [zoom]);

  useEffect(() => {
    const api = apiRef.current;
    if (!api) return;

    // Give wrap/rotation rows a bit more room.
    api.setRowHeight(16, 60);
    api.setRowHeight(17, 60);
    api.setRowHeight(20, 48);

    // Slightly wider first few columns so labels are readable.
    api.setColWidth(1, 160);
    api.setColWidth(2, 220);
    api.setColWidth(3, 180);
  }, []);

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
        <div style={{ fontWeight: 600 }}>Cell formatting demo (Excel-style rendering)</div>
        <div style={{ marginTop: 6, color: "var(--formula-grid-cell-text, #4b5563)", opacity: 0.8 }}>
          Use this view to manually verify cell formatting changes in <code>@formula/grid</code>.
        </div>

        <div style={{ marginTop: 10, display: "flex", gap: 18, flexWrap: "wrap", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>Legend</div>
            <ul style={{ margin: 0, paddingLeft: 18, lineHeight: 1.4, opacity: 0.9 }}>
              <li>
                <code>A1:D4</code>: bold / italic / underline / strike combinations
              </li>
              <li>
                <code>A5:C6</code>: fills + explicit font color / font family
              </li>
              <li>
                <code>A8:C13</code>: borders (thin/medium/thick, dashed/dotted/double) + conflicts
              </li>
              <li>
                <code>A14:C17</code>: alignment, wrap on/off, rotation
              </li>
              <li>
                <code>A20</code>: rich text runs (if Task 84 lands)
              </li>
              <li>
                <code>A22:C23</code>: number format display strings (Task 68)
              </li>
            </ul>
          </div>

          <div style={{ minWidth: 240 }}>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>Switch demos</div>
            <div style={{ opacity: 0.9, lineHeight: 1.4 }}>
              <div>
                <code>?demo=style</code> (this)
              </div>
              <div>
                <code>?demo=perf</code> (performance harness)
              </div>
              <div>
                <code>?demo=merged</code> (merged cells demo)
              </div>
            </div>
          </div>

          <div style={{ flex: "1 1 240px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8 }}>
              Zoom
              <input
                type="range"
                min={0.5}
                max={3}
                step={0.1}
                value={zoom}
                onChange={(event) => setZoom(event.currentTarget.valueAsNumber)}
              />
              <span style={{ width: 44, textAlign: "right" }}>{Math.round(zoom * 100)}%</span>
            </label>
          </div>
        </div>
      </div>

      <div style={{ flex: 1, position: "relative" }}>
        <CanvasGrid
          provider={provider}
          rowCount={rowCount}
          colCount={colCount}
          headerRows={1}
          headerCols={1}
          frozenRows={1}
          frozenCols={1}
          defaultRowHeight={24}
          defaultColWidth={120}
          onZoomChange={setZoom}
          apiRef={apiRef}
        />
      </div>
    </div>
  );
}
