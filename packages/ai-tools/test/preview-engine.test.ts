import { describe, expect, it, vi } from "vitest";
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

  it("requires approval for fetch_external_data previews", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine();

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "fetch_external_data",
          parameters: {
            source_type: "api",
            url: "https://example.com/data",
            destination: "Sheet1!A1"
          }
        }
      ],
      workbook
    );

    expect(preview.requires_approval).toBe(true);
    expect(preview.approval_reasons).toContain("External data access requested");
  });

  it("never performs network access during preview, even when executor options enable it", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const previewEngine = new PreviewEngine();

    const fetchMock = vi.fn(async () => {
      throw new Error("fetch should not be called during preview");
    });
    vi.stubGlobal("fetch", fetchMock as any);

    const preview = await previewEngine.generatePreview(
      [
        {
          name: "fetch_external_data",
          parameters: {
            source_type: "api",
            url: "https://example.com/data",
            destination: "Sheet1!A1"
          }
        }
      ],
      workbook,
      { allow_external_data: true, allowed_external_hosts: ["example.com"] }
    );

    expect(fetchMock).not.toHaveBeenCalled();
    expect(preview.requires_approval).toBe(true);
    expect(preview.tool_results[0]?.ok).toBe(false);
    expect(preview.tool_results[0]?.error?.code).toBe("permission_denied");
  });
});
