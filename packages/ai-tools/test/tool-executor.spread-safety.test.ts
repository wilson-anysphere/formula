import { describe, expect, it } from "vitest";

import { ToolExecutor } from "../src/executor/tool-executor.js";
import { columnIndexToLabel } from "../src/spreadsheet/a1.js";

describe("ToolExecutor (spread safety)", () => {
  it("compute_statistics min/max handles large ranges without spread argument limits", async () => {
    const cols = 200_000;
    const endCol = columnIndexToLabel(cols);
    const range = `Sheet1!A1:${endCol}1`;

    // Avoid allocating 200k CellData objects up front; ToolExecutor only reads each cell once.
    const cell: any = { value: 0 };
    const row = new Proxy(
      { length: cols } as any,
      {
        get(_target, prop) {
          if (prop === "length") return cols;
          const idx = Number(prop);
          if (Number.isInteger(idx) && idx >= 0 && idx < cols) {
            cell.value = idx;
            return cell;
          }
          return undefined;
        },
      },
    );

    const spreadsheet: any = {
      readRange() {
        return [row];
      },
    };

    const executor = new ToolExecutor(spreadsheet);
    const result = await executor.execute({
      name: "compute_statistics",
      parameters: { range, measures: ["min", "max"] },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("compute_statistics");
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

    expect(result.data?.statistics.min).toBe(0);
    expect(result.data?.statistics.max).toBe(cols - 1);
  });

  it("compute_statistics does not materialize values[] for streaming-only measures", async () => {
    const cols = 200_000;
    const endCol = columnIndexToLabel(cols);
    const range = `Sheet1!A1:${endCol}1`;

    // Avoid allocating 200k CellData objects up front; ToolExecutor only reads each cell once.
    const cell: any = { value: 0 };
    const row = new Proxy(
      { length: cols } as any,
      {
        get(_target, prop) {
          if (prop === "length") return cols;
          const idx = Number(prop);
          if (Number.isInteger(idx) && idx >= 0 && idx < cols) {
            cell.value = idx;
            return cell;
          }
          return undefined;
        },
      },
    );

    const spreadsheet: any = {
      readRange() {
        return [row];
      },
    };

    const originalPush = Array.prototype.push;
    let pushCount = 0;
    // Avoid vi.spyOn(Array.prototype, "push") here: spy implementations often use push
    // internally to record calls, which can recurse when the method being spied is `push`.
    (Array.prototype as any).push = function (...args: any[]) {
      pushCount++;
      return originalPush.apply(this, args as any);
    };
    try {
      const executor = new ToolExecutor(spreadsheet);
      const result = await executor.execute({
        name: "compute_statistics",
        parameters: { range, measures: ["min", "max", "mean"] },
      });

      expect(result.ok).toBe(true);
      expect(result.tool).toBe("compute_statistics");
      if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

      expect(result.data?.statistics.min).toBe(0);
      expect(result.data?.statistics.max).toBe(cols - 1);
      expect(result.data?.statistics.mean).toBeCloseTo((cols - 1) / 2, 8);

      // Regression guard: streaming-only measures should not build a `values[]` list by
      // pushing each numeric cell. (Allow a small number of incidental pushes elsewhere.)
      expect(pushCount).toBeLessThan(1_000);
    } finally {
      Array.prototype.push = originalPush;
    }
  });

  it("compute_statistics correlation handles large ranges without materializing values[]", async () => {
    const rows = 100_000; // 2 * 100k = 200k cells (default max_tool_range_cells)
    const range = `Sheet1!A1:B${rows}`;

    const leftCell: any = { value: 0 };
    const rightCell: any = { value: 0 };
    let currentRow = 0;

    const row = new Proxy(
      { length: 2 } as any,
      {
        get(_target, prop) {
          if (prop === "length") return 2;
          const idx = Number(prop);
          if (idx === 0) {
            leftCell.value = currentRow;
            return leftCell;
          }
          if (idx === 1) {
            rightCell.value = currentRow * 2;
            return rightCell;
          }
          return undefined;
        },
      },
    );

    const cells = new Proxy(
      { length: rows } as any,
      {
        get(_target, prop) {
          if (prop === "length") return rows;
          const idx = Number(prop);
          if (Number.isInteger(idx) && idx >= 0 && idx < rows) {
            currentRow = idx;
            return row;
          }
          return undefined;
        },
      },
    );

    const spreadsheet: any = {
      readRange() {
        return cells;
      },
    };

    const originalPush = Array.prototype.push;
    let pushCount = 0;
    (Array.prototype as any).push = function (...args: any[]) {
      pushCount++;
      return originalPush.apply(this, args as any);
    };
    try {
      const executor = new ToolExecutor(spreadsheet);
      const result = await executor.execute({
        name: "compute_statistics",
        parameters: { range, measures: ["correlation"] },
      });

      expect(result.ok).toBe(true);
      expect(result.tool).toBe("compute_statistics");
      if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");

      expect(result.data?.statistics.correlation).toBeCloseTo(1, 12);
      // No per-cell `values[]` pushes should occur.
      expect(pushCount).toBeLessThan(1_000);
    } finally {
      Array.prototype.push = originalPush;
    }
  });
});
