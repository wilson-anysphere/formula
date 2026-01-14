import { describe, expect, it } from "vitest";

import { ToolExecutor } from "./tool-executor.ts";
import { parseA1Cell } from "../spreadsheet/a1.ts";
import { InMemoryWorkbook } from "../spreadsheet/in-memory-workbook.ts";

import { DLP_ACTION } from "../../../security/dlp/src/actions.js";

describe("ToolExecutor read_range DLP + include_formula_values", () => {
  it("does not redact formulas based on computed formula values when DLP is already REDACT", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    // Range includes an explicitly restricted cell (B1) so the range-level DLP decision is REDACT.
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "secret" });
    // Simulate a real backend providing a cached/computed formula value alongside the formula text.
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "secret", formula: "=B1" });

    const executor = new ToolExecutor(workbook, {
      default_sheet: "Sheet1",
      include_formula_values: true,
      dlp: {
        document_id: "doc-1",
        policy: {
          version: 1,
          allowDocumentOverrides: true,
          rules: {
            [DLP_ACTION.AI_CLOUD_PROCESSING]: {
              maxAllowed: "Internal",
              allowRestrictedContent: false,
              redactDisallowed: true,
            },
          },
        },
        classification_records: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 }, // B1
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      parameters: { range: "Sheet1!A1:B1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    expect(result.tool).toBe("read_range");
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    // Under DLP REDACT, formula values should remain null even when include_formula_values is enabled.
    // (Formula values can depend on restricted dependencies, and ToolExecutor does not trace provenance.)
    expect(result.data?.values).toEqual([[null, "[REDACTED]"]]);
    // The formula text itself should remain visible for allowed cells; it should not be redacted just
    // because the backend provided a cached/computed formula value.
    expect(result.data?.formulas).toEqual([["=B1", "[REDACTED]"]]);
  });
});

