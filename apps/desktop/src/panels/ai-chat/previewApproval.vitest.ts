import { describe, expect, it } from "vitest";

import { formatPreviewApprovalPrompt } from "./previewApproval.js";

describe("formatPreviewApprovalPrompt", () => {
  it("includes summary + a few cell diffs", () => {
    const prompt = formatPreviewApprovalPrompt(
      {
        call: { name: "write_cell", arguments: { cell: "Sheet1!A1", value: 123 } },
        preview: {
          timing_ms: 5,
          tool_results: [],
          changes: [
            {
              cell: "Sheet1!A1",
              type: "modify",
              before: { value: null },
              after: { value: 123 },
            },
          ],
          summary: { total_changes: 1, creates: 0, modifies: 1, deletes: 0 },
          warnings: ["unit-test warning"],
          requires_approval: true,
          approval_reasons: ["Large edit (1 cells)"],
        },
      },
      { max_changes: 5 },
    );

    expect(prompt).toContain("AI wants to run: write_cell");
    expect(prompt).toContain("Summary: 1 changes");
    expect(prompt).toContain("unit-test warning");
    expect(prompt).toContain("Sheet1!A1");
  });
});

