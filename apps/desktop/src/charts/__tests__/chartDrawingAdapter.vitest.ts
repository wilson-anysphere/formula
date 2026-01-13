import { describe, expect, it } from "vitest";

import { anchorToRectPx, pxToEmu } from "../../drawings/overlay";
import type { GridGeometry } from "../../drawings/overlay";
import type { ChartRecord } from "../chartStore";
import { chartRecordToDrawingObject } from "../chartDrawingAdapter";

const geom: GridGeometry = {
  cellOriginPx: ({ row, col }) => ({ x: col * 100, y: row * 20 }),
  cellSizePx: () => ({ width: 100, height: 20 }),
};

describe("charts/chartDrawingAdapter", () => {
  it("maps a twoCell chart anchor into a DrawingObject with a correct rect", () => {
    const chart: ChartRecord = {
      id: "chart_1",
      sheetId: "Sheet1",
      chartType: { kind: "bar" as any },
      title: "Test",
      series: [],
      anchor: {
        kind: "twoCell",
        fromCol: 2,
        fromRow: 1,
        fromColOffEmu: 0,
        fromRowOffEmu: 0,
        toCol: 5,
        toRow: 4,
        toColOffEmu: 0,
        toRowOffEmu: 0,
      },
    };

    const obj = chartRecordToDrawingObject(chart);
    expect(obj.kind.type).toBe("chart");
    if (obj.kind.type !== "chart") {
      throw new Error(`Expected chart DrawingObjectKind, got ${obj.kind.type}`);
    }
    expect(obj.kind.chartId).toBe("chart_1");

    const rect = anchorToRectPx(obj.anchor, geom);
    expect(rect).toEqual({ x: 200, y: 20, width: 300, height: 60 });
  });

  it("maps an absolute anchor and preserves EMU sizes", () => {
    const chart: ChartRecord = {
      id: "chart_2",
      sheetId: "Sheet1",
      chartType: { kind: "line" as any },
      series: [],
      anchor: {
        kind: "absolute",
        xEmu: pxToEmu(10),
        yEmu: pxToEmu(12),
        cxEmu: pxToEmu(30),
        cyEmu: pxToEmu(40),
      },
    };

    const obj = chartRecordToDrawingObject(chart);
    const rect = anchorToRectPx(obj.anchor, geom);
    expect(rect).toEqual({ x: 10, y: 12, width: 30, height: 40 });
  });
});
