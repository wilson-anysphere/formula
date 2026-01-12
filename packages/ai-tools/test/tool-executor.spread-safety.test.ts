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
});

