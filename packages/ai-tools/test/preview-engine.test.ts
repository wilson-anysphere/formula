import { describe, expect, it } from "vitest";
import { PreviewEngine } from "../src/preview/preview-engine.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";

describe("PreviewEngine", () => {
  it("flags large edits for approval and truncates change list", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine({ max_preview_changes: 20, approval_cell_threshold: 100 });

    const values = Array.from({ length: 10 }, (_, r) =>
      Array.from({ length: 15 }, (_, c) => `${r + 1}:${c + 1}`)
    );

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "set_range",
          parameters: {
            range: "Sheet1!A1:O10",
            values
          }
        }
      ],
      workbook
    );

    expect(preview.summary.total_changes).toBe(150);
    expect(preview.changes.length).toBe(20);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.startsWith("Large edit"))).toBe(true);
  });

  it("detects deletes and requires approval", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 123 });

    const previewEngine = new PreviewEngine();
    const preview = await previewEngine.generatePreview(
      [
        {
          name: "write_cell",
          parameters: { cell: "Sheet1!A1", value: null }
        }
      ],
      workbook
    );

    expect(preview.summary.deletes).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons.some((reason) => reason.includes("Deletes"))).toBe(true);
  });
});

