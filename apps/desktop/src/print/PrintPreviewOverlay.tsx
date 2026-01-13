import * as React from "react";

import { calculatePages } from "./pageBreaks";
import type { CellRange, ManualPageBreaks, PageSetup } from "./types";

type Props = {
  printArea: CellRange;
  pageSetup: PageSetup;
  manualBreaks?: ManualPageBreaks;
  colWidthsPoints: number[];
  rowHeightsPoints: number[];
};

function sumBefore(values: number[], idx1: number): number {
  let sum = 0;
  for (let i = 1; i < idx1; i++) sum += values[i - 1] ?? 0;
  return sum;
}

function sumRange(values: number[], start1: number, end1: number): number {
  let sum = 0;
  for (let i = start1; i <= end1; i++) sum += values[i - 1] ?? 0;
  return sum;
}

export function PrintPreviewOverlay({
  printArea,
  pageSetup,
  manualBreaks,
  colWidthsPoints,
  rowHeightsPoints,
}: Props) {
  const pages = React.useMemo(
    () =>
      calculatePages(printArea, colWidthsPoints, rowHeightsPoints, pageSetup, manualBreaks),
    [printArea, colWidthsPoints, rowHeightsPoints, pageSetup, manualBreaks],
  );

  // Coordinates are in “grid points” (often close to CSS px at 96dpi) and should be transformed
  // by the caller into screen space if the grid is zoomed/scrolled.
  const rects = React.useMemo(() => {
    return pages.map((p) => {
      const x = sumBefore(colWidthsPoints, p.startCol);
      const y = sumBefore(rowHeightsPoints, p.startRow);
      const w = sumRange(colWidthsPoints, p.startCol, p.endCol);
      const h = sumRange(rowHeightsPoints, p.startRow, p.endRow);
      return { x, y, w, h };
    });
  }, [pages, colWidthsPoints, rowHeightsPoints]);

  const width = sumRange(colWidthsPoints, 1, colWidthsPoints.length);
  const height = sumRange(rowHeightsPoints, 1, rowHeightsPoints.length);

  return (
    <svg
      width={width}
      height={height}
      className="print-preview-overlay"
    >
      {rects.map((r, idx) => (
        <rect
          key={idx}
          x={r.x}
          y={r.y}
          width={r.w}
          height={r.h}
          fill="none"
          stroke="var(--border)"
          strokeOpacity={0.35}
          strokeWidth={1}
          strokeDasharray="6 4"
        />
      ))}
    </svg>
  );
}
