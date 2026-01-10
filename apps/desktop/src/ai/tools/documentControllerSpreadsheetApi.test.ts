import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { ToolExecutor, PreviewEngine } from "../../../../../packages/ai-tools/src/index.js";

import { DocumentControllerSpreadsheetApi } from "./documentControllerSpreadsheetApi.js";

describe("DocumentControllerSpreadsheetApi", () => {
  it("allows ToolExecutor to apply updates through DocumentController", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const api = new DocumentControllerSpreadsheetApi(controller);
    const executor = new ToolExecutor(api, { default_sheet: "Sheet1" });

    const result = await executor.execute({
      name: "write_cell",
      parameters: { cell: "A1", value: 99 }
    });

    expect(result.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(99);
  });

  it("supports PreviewEngine diffing without mutating the live controller", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 10);
    controller.setCellValue("Sheet1", "B1", { text: "Rich Bold", runs: [{ start: 0, end: 4, style: { bold: true } }] });

    const api = new DocumentControllerSpreadsheetApi(controller);
    const previewEngine = new PreviewEngine({ approval_cell_threshold: 0 });

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "write_cell",
          parameters: { cell: "A1", value: 20 }
        }
      ],
      api,
      { default_sheet: "Sheet1" }
    );

    expect(preview.summary.total_changes).toBe(1);
    expect(preview.summary.modifies).toBe(1);
    expect(preview.requires_approval).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(10);

    // Sanity check: tool call references normalize to Sheet1.
    expect(preview.changes[0]?.cell).toBe("Sheet1!A1");
  });
});
