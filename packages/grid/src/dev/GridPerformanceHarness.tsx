import React, { useEffect, useMemo, useRef } from "react";
import { CanvasGrid, type GridApi } from "../react/CanvasGrid";
import { MockCellProvider } from "../model/MockCellProvider";

export function GridPerformanceHarness(props?: {
  rowCount?: number;
  colCount?: number;
  frames?: number;
  deltaY?: number;
}): React.ReactElement {
  const rowCount = props?.rowCount ?? 1_000_000;
  const colCount = props?.colCount ?? 100;
  const frames = props?.frames ?? 180;
  const deltaY = props?.deltaY ?? 120;

  const apiRef = useRef<GridApi | null>(null);

  const provider = useMemo(() => new MockCellProvider({ rowCount, colCount }), [rowCount, colCount]);

  useEffect(() => {
    const nodeEnv = (globalThis as any)?.process?.env?.NODE_ENV as string | undefined;
    if (nodeEnv === "production") return;
    const api = apiRef.current;
    if (!api) return;

    let remaining = frames;
    let last = performance.now();
    const samples: number[] = [];

    const tick = (now: number) => {
      const dt = now - last;
      last = now;
      samples.push(dt);

      api.scrollBy(0, deltaY);
      remaining -= 1;
      if (remaining > 0) {
        requestAnimationFrame(tick);
      } else {
        const trimmed = samples.slice(1);
        const avg = trimmed.reduce((sum, value) => sum + value, 0) / Math.max(1, trimmed.length);
        console.log(
          `[grid-perf] frames=${trimmed.length} avgFrame=${avg.toFixed(2)}ms (~${(1000 / avg).toFixed(1)}fps)`
        );
      }
    };

    requestAnimationFrame((now) => {
      last = now;
      requestAnimationFrame(tick);
    });
  }, [frames, deltaY]);

  return (
    <div style={{ width: "100%", height: "100%", position: "relative" }}>
      <CanvasGrid provider={provider} rowCount={rowCount} colCount={colCount} frozenRows={1} frozenCols={1} apiRef={apiRef} />
    </div>
  );
}
