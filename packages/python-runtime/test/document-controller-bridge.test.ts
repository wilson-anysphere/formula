import { describe, expect, it } from "vitest";

import { DocumentController } from "../../../apps/desktop/src/document/documentController.js";
import { NativePythonRuntime } from "@formula/python-runtime/native";
import { DocumentControllerBridge } from "@formula/python-runtime/document-controller";

describe("DocumentControllerBridge", () => {
  it("lets Python scripts write values + formulas into a DocumentController", async () => {
    const doc = new DocumentController();
    const api = new DocumentControllerBridge(doc, { activeSheetId: "Sheet1" });
    const runtime = new NativePythonRuntime({
      timeoutMs: 10_000,
      maxMemoryBytes: 256 * 1024 * 1024,
      permissions: { filesystem: "none", network: "none" },
    });

    const script = `
import formula

sheet = formula.active_sheet
sheet["A1"] = 42
sheet["A2"] = "=A1*2"
`;

    await runtime.execute(script, { api });

    expect(doc.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(42);
    expect(doc.getCell("Sheet1", { row: 1, col: 0 }).formula).toBe("=A1*2");
  });
});
