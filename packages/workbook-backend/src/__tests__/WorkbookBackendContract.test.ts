import { describe, expect, it } from "vitest";

import { WORKBOOK_BACKEND_REQUIRED_METHODS } from "../index";
import { WasmWorkbookBackend } from "../../../engine/src/backend/WasmWorkbookBackend";
import { TauriWorkbookBackend } from "../../../../apps/desktop/src/tauri/workbookBackend";

function expectImplementsWorkbookBackendContract(ctor: { prototype: Record<string, unknown> }): void {
  for (const method of WORKBOOK_BACKEND_REQUIRED_METHODS) {
    expect(typeof ctor.prototype[method]).toBe("function");
  }
}

describe("WorkbookBackend contract", () => {
  it("WasmWorkbookBackend exposes the required methods", () => {
    expectImplementsWorkbookBackendContract(WasmWorkbookBackend);
  });

  it("TauriWorkbookBackend exposes the required methods", () => {
    expectImplementsWorkbookBackendContract(TauriWorkbookBackend);
  });
});

