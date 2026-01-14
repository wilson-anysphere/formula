import type { Anchor as DrawingAnchor, DrawingObject } from "../drawings/types";
import type { ChartAnchor, ChartRecord } from "./chartStore";

function stableHash32(input: string): number {
  // FNV-1a 32-bit.
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  // Ensure positive safe integer.
  return hash >>> 0;
}

export function chartIdToDrawingId(chartId: string): number {
  // Drawing object ids are globally unique numbers (often random 53-bit). To avoid collisions
  // when charts are rendered alongside workbook drawings, keep chart ids in a separate
  // namespace by using negative numbers derived from a stable hash.
  //
  // Note: we intentionally avoid reusing the numeric suffix from `chart_42` ids because
  // workbook drawings frequently contain small numeric ids as well.
  const hashed = stableHash32(String(chartId ?? ""));
  // Avoid `0` which can be treated as a sentinel in some callers.
  //
  // When the hash is 0, use the reserved `-2^32` value so we stay within the chart namespace
  // without colliding with `hashed === 1` (which would otherwise also map to `-1`).
  return hashed === 0 ? -0x100000000 : -hashed;
}

/**
 * Returns true when the provided drawing object id belongs to a ChartStore (canvas) chart.
 *
 * ChartStore charts use stable negative 32-bit ids (see `chartIdToDrawingId`). Workbook drawings
 * may also use negative ids when their raw ids are not JS-safe; those hashed ids live in a
 * separate, large-magnitude negative namespace (see `parseDrawingObjectId` in
 * `drawings/modelAdapters.ts`).
 */
export function isChartStoreDrawingId(id: number): boolean {
  // `chartIdToDrawingId` returns values in `[-2^32, -1]`.
  //
  // Workbook drawings can also use negative ids when their raw ids are not JS-safe; those hashed
  // ids live at `<= -2^33` (see `parseDrawingObjectId` in `drawings/modelAdapters.ts`) so they do
  // not collide with the chart namespace.
  return typeof id === "number" && Number.isFinite(id) && id < 0 && id >= -0x100000000;
}

export function chartAnchorToDrawingAnchor(anchor: ChartAnchor): DrawingAnchor {
  switch (anchor.kind) {
    case "absolute":
      return {
        type: "absolute",
        pos: { xEmu: anchor.xEmu, yEmu: anchor.yEmu },
        size: { cx: anchor.cxEmu, cy: anchor.cyEmu },
      };
    case "oneCell":
      return {
        type: "oneCell",
        from: {
          cell: { row: anchor.fromRow, col: anchor.fromCol },
          offset: { xEmu: anchor.fromColOffEmu, yEmu: anchor.fromRowOffEmu },
        },
        size: { cx: anchor.cxEmu, cy: anchor.cyEmu },
      };
    case "twoCell":
      return {
        type: "twoCell",
        from: {
          cell: { row: anchor.fromRow, col: anchor.fromCol },
          offset: { xEmu: anchor.fromColOffEmu, yEmu: anchor.fromRowOffEmu },
        },
        to: {
          cell: { row: anchor.toRow, col: anchor.toCol },
          offset: { xEmu: anchor.toColOffEmu, yEmu: anchor.toRowOffEmu },
        },
      };
  }
}

export function drawingAnchorToChartAnchor(anchor: DrawingAnchor): ChartAnchor {
  switch (anchor.type) {
    case "absolute":
      return {
        kind: "absolute",
        xEmu: anchor.pos.xEmu,
        yEmu: anchor.pos.yEmu,
        cxEmu: anchor.size.cx,
        cyEmu: anchor.size.cy,
      };
    case "oneCell":
      return {
        kind: "oneCell",
        fromCol: anchor.from.cell.col,
        fromRow: anchor.from.cell.row,
        fromColOffEmu: anchor.from.offset.xEmu,
        fromRowOffEmu: anchor.from.offset.yEmu,
        cxEmu: anchor.size.cx,
        cyEmu: anchor.size.cy,
      };
    case "twoCell":
      return {
        kind: "twoCell",
        fromCol: anchor.from.cell.col,
        fromRow: anchor.from.cell.row,
        fromColOffEmu: anchor.from.offset.xEmu,
        fromRowOffEmu: anchor.from.offset.yEmu,
        toCol: anchor.to.cell.col,
        toRow: anchor.to.cell.row,
        toColOffEmu: anchor.to.offset.xEmu,
        toRowOffEmu: anchor.to.offset.yEmu,
      };
  }
}

export function chartRecordToDrawingObject(chart: ChartRecord, zOrder: number = 0): DrawingObject {
  return {
    id: chartIdToDrawingId(chart.id),
    kind: { type: "chart", chartId: chart.id, label: chart.title },
    anchor: chartAnchorToDrawingAnchor(chart.anchor),
    zOrder,
    ...(chart.anchor.kind === "absolute" || chart.anchor.kind === "oneCell"
      ? { size: { cx: chart.anchor.cxEmu, cy: chart.anchor.cyEmu } }
      : {}),
  };
}

export function chartRecordsToDrawingObjects(charts: readonly ChartRecord[]): DrawingObject[] {
  return charts.map((chart, idx) => chartRecordToDrawingObject(chart, idx));
}
