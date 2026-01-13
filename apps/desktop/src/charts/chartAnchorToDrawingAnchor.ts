import type { Anchor } from "../drawings/types";

import type { ChartAnchor } from "./chartStore";

export function chartAnchorToDrawingAnchor(anchor: ChartAnchor): Anchor {
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

