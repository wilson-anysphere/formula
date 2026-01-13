import React, { useEffect, useMemo, useRef, useState } from "react";

import type {
  CellBorderLineStyle,
  CellBorderSpec,
  CellBorders,
  CellData,
  CellRange,
  CellProvider,
  CellStyle
} from "../model/CellProvider";
import { toA1Address, toColumnName } from "../a11y/a11y";
import { CanvasGrid, type GridApi } from "../react/CanvasGrid";

function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.startRow < b.endRow && a.endRow > b.startRow && a.startCol < b.endCol && a.endCol > b.startCol;
}

function rangeContains(range: CellRange, row: number, col: number): boolean {
  return row >= range.startRow && row < range.endRow && col >= range.startCol && col < range.endCol;
}

class CellFormattingDemoProvider implements CellProvider {
  private readonly rowCount: number;
  private readonly colCount: number;
  private readonly cells = new Map<string, CellData>();
  private readonly mergedRanges: CellRange[] = [];

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

    put(2, 5, "Double underline", {
      ...baseTextStyle,
      underline: true,
      underlineStyle: "double"
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
    // Font sizes / weights
    // ---------------------------------------------------------------------------------------------
    put(7, 1, "Font size / weight", { fill: sectionHeaderFill, fontWeight: "700" });
    put(7, 2, "10px", { fontSize: 10, textAlign: "center", fill: "rgba(148, 163, 184, 0.08)" });
    put(7, 3, "20px bold", { fontSize: 20, fontWeight: "700", textAlign: "center", fill: "rgba(148, 163, 184, 0.08)" });
    put(7, 4, "900 weight", { fontWeight: "900", textAlign: "center", fill: "rgba(148, 163, 184, 0.08)" });

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

    // Merged border example (anchor at E11, spans E11:F12).
    const mergedBorderRange: CellRange = { startRow: 11, endRow: 13, startCol: 5, endCol: 7 };
    this.mergedRanges.push(mergedBorderRange);
    put(mergedBorderRange.startRow, mergedBorderRange.startCol, "Merged border\n(E11:F12)", {
      ...borderCellBase,
      fill: "rgba(168, 85, 247, 0.12)",
      fontWeight: "700",
      wrapMode: "word",
      borders: allBorders(mkBorder(3, "double", "#a855f7"))
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

    put(11, 3, "Mixed edges", {
      ...borderCellBase,
      borders: {
        top: mkBorder(3, "solid", "#ef4444"),
        right: mkBorder(2, "dashed", "#22c55e"),
        bottom: mkBorder(2, "double", "#a855f7"),
        left: mkBorder(1, "dotted", "#3b82f6")
      }
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

    // Wrap disabled: text should overflow into adjacent empty cells until blocked.
    put(16, 1, "Wrap off: overflow into empty cells →", { wrapMode: "none", textAlign: "start" });
    put(16, 4, "STOP", { fill: "#fee2e2", fontWeight: "700", textAlign: "center" });

    // Wrap enabled: long text should wrap inside the cell.
    put(
      17,
      1,
      "Wrap on: This long text should wrap inside the cell (word wrap) instead of overflowing into neighbors.",
      { wrapMode: "word", textAlign: "start" }
    );

    put(17, 3, "Rotated 45°", {
      rotationDeg: 45,
      textAlign: "center",
      verticalAlign: "middle",
      fill: "rgba(14, 101, 235, 0.12)",
      fontWeight: "600"
    });

    put(17, 4, "Rotated 90°", {
      rotationDeg: 90,
      textAlign: "center",
      verticalAlign: "middle",
      fill: "rgba(14, 101, 235, 0.12)",
      fontWeight: "600"
    });

    // Vertical alignment (top/middle/bottom).
    put(18, 1, "Top\naligned", { verticalAlign: "top", textAlign: "center", fill: "rgba(148, 163, 184, 0.06)" });
    put(18, 2, "Middle\naligned", { verticalAlign: "middle", textAlign: "center", fill: "rgba(148, 163, 184, 0.06)" });
    put(18, 3, "Bottom\naligned", { verticalAlign: "bottom", textAlign: "center", fill: "rgba(148, 163, 184, 0.06)" });

    // Char wrap + RTL/LTR direction.
    put(19, 1, "CharWrap: 1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", {
      wrapMode: "char",
      textAlign: "start"
    });
    put(19, 2, "Auto RTL (start-align): שלום world 123", {
      wrapMode: "word",
      textAlign: "start",
      direction: "auto"
    });
    put(19, 3, "RTL (end-align): שלום world 123", {
      wrapMode: "word",
      textAlign: "end",
      direction: "rtl"
    });

    // ---------------------------------------------------------------------------------------------
    // Rich text runs
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
  const [selection, setSelection] = useState<{ row: number; col: number } | null>(null);
  const [selectionRange, setSelectionRange] = useState<CellRange | null>(null);

  const selectionCell = selection ? provider.getCell(selection.row, selection.col) : null;
  const selectionAddress = (() => {
    if (!selection) return "None";
    const headerRows = 1;
    const headerCols = 1;
    const row0 = selection.row - headerRows;
    const col0 = selection.col - headerCols;
    if (row0 >= 0 && col0 >= 0) return toA1Address(row0, col0);
    return `R${selection.row + 1}C${selection.col + 1}`;
  })();

  const selectionRangeA1 = (() => {
    if (!selectionRange) return null;
    const headerRows = 1;
    const headerCols = 1;
    const startRow0 = selectionRange.startRow - headerRows;
    const startCol0 = selectionRange.startCol - headerCols;
    const endRow0 = selectionRange.endRow - headerRows - 1;
    const endCol0 = selectionRange.endCol - headerCols - 1;
    if (startRow0 < 0 || startCol0 < 0 || endRow0 < 0 || endCol0 < 0) return null;
    const start = toA1Address(startRow0, startCol0);
    const end = toA1Address(endRow0, endCol0);
    return start === end ? start : `${start}:${end}`;
  })();

  useEffect(() => {
    apiRef.current?.setZoom(zoom);
  }, [zoom]);

  useEffect(() => {
    const api = apiRef.current;
    if (!api) return;

    // Give wrap/rotation rows a bit more room.
    api.setRowHeight(7, 34);
    api.setRowHeight(16, 60);
    api.setRowHeight(17, 60);
    api.setRowHeight(18, 68);
    api.setRowHeight(19, 60);
    api.setRowHeight(20, 48);

    // Slightly wider first few columns so labels are readable.
    api.setColWidth(1, 160);
    api.setColWidth(2, 220);
    api.setColWidth(3, 180);
    api.setColWidth(4, 140);
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
        <div style={{ marginTop: 6, color: "var(--formula-grid-cell-text, #4b5563)", opacity: 0.7, fontSize: 12 }}>
          Tip: drag row/column header boundaries to resize; double-click to auto-fit.
        </div>

        <div style={{ marginTop: 10, display: "flex", gap: 18, flexWrap: "wrap", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>Legend</div>
            <ul style={{ margin: 0, paddingLeft: 18, lineHeight: 1.4, opacity: 0.9 }}>
              <li>
                <code>A1:E4</code>: bold / italic / underline / strike / double underline combinations
              </li>
              <li>
                <code>A5:C6</code>: fills + explicit font color / font family
              </li>
              <li>
                <code>A7:D7</code>: font size / weight
              </li>
              <li>
                <code>A8:C13</code>: borders (thin/medium/thick, dashed/dotted/double) + conflicts
              </li>
              <li>
                <code>E11:F12</code>: merged cell borders
              </li>
              <li>
                <code>A14:D17</code>: alignment, wrap on/off, rotation
              </li>
              <li>
                <code>A18:C19</code>: vertical alignment + char wrap + RTL/LTR direction
              </li>
              <li>
                <code>A20</code>: rich text runs
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
                <a href="?demo=style" style={{ color: "inherit", textDecoration: "none" }}>
                  <code>?demo=style</code>
                </a>{" "}
                (this)
              </div>
              <div>
                <a href="?demo=perf" style={{ color: "inherit", textDecoration: "none" }}>
                  <code>?demo=perf</code>
                </a>{" "}
                (performance harness)
              </div>
              <div>
                <a href="?demo=merged" style={{ color: "inherit", textDecoration: "none" }}>
                  <code>?demo=merged</code>
                </a>{" "}
                (merged cells demo)
              </div>
            </div>
          </div>

          <div style={{ flex: "1 1 240px" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8 }}>
              Zoom
                <input
                  type="range"
                  min={0.25}
                  max={4}
                  step={0.05}
                  value={zoom}
                  onChange={(event) => setZoom(event.currentTarget.valueAsNumber)}
                />
              <span style={{ width: 44, textAlign: "right" }}>{Math.round(zoom * 100)}%</span>
            </label>

            <details style={{ marginTop: 10, opacity: 0.95 }}>
              <summary style={{ cursor: "pointer", userSelect: "none" }}>Selection inspector</summary>
              <div style={{ marginTop: 8, fontSize: 12, lineHeight: 1.4 }}>
                <div>
                  <strong>Active:</strong> {selectionAddress}
                </div>
                {selectionCell ? (
                  <pre
                    style={{
                      marginTop: 6,
                      marginBottom: 0,
                      padding: 8,
                      borderRadius: 6,
                      border: "1px solid var(--formula-grid-line, #e5e7eb)",
                      background: "rgba(148, 163, 184, 0.08)",
                      overflow: "auto",
                      maxHeight: 160
                    }}
                  >
                    {JSON.stringify(
                      {
                        value: selectionCell.value,
                        style: selectionCell.style,
                        richText: selectionCell.richText
                      },
                      null,
                      2
                    )}
                  </pre>
                ) : (
                  <div style={{ marginTop: 6, opacity: 0.8 }}>Select a cell to see its value/style payload.</div>
                )}
                {selectionRange ? (
                  <div style={{ marginTop: 6 }}>
                    <strong>Range:</strong>{" "}
                    <code>{selectionRangeA1 ?? "(includes headers)"}</code>
                  </div>
                ) : null}
              </div>
            </details>
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
          onSelectionChange={setSelection}
          onSelectionRangeChange={setSelectionRange}
          enableResize
          apiRef={apiRef}
        />
      </div>
    </div>
  );
}
