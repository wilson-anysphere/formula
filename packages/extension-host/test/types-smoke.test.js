const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");

test("extension-api types file exports the expected surface", async () => {
  const dtsPath = path.resolve(__dirname, "../../extension-api/index.d.ts");
  const text = await require("node:fs/promises").readFile(dtsPath, "utf8");
  const { stripComments } = await import("../../../apps/desktop/test/sourceTextUtils.js");
  const source = stripComments(text);

  // Smoke checks only: the repo CI may not run a full TS compile, but we still want to catch
  // accidental removals of key API declarations promised by docs/10-extensibility.md.
  for (const fragment of [
    "export namespace cells",
    "function getRange(ref: string): Promise<Range>;",
    "function setRange(ref: string, values: CellValue[][]): Promise<void>;",
    "readonly sheets: Sheet[];",
    "readonly activeSheet: Sheet;",
    "readonly address: string;",
    "readonly formulas: (string | null)[][];",
    "export namespace sheets",
    "function createSheet(name: string): Promise<Sheet>;",
    "function deleteSheet(name: string): Promise<void>;",
    "getRange(ref: string): Promise<Range>;",
    "setRange(ref: string, values: CellValue[][]): Promise<void>;",
    "export namespace events",
    "function onSheetActivated",
    "function onWorkbookOpened",
    "function onBeforeSave",
    "export namespace dataConnectors",
    "function register(connectorId: string, impl: DataConnectorImplementation): Promise<Disposable>;"
  ]) {
    assert.ok(source.includes(fragment), `Missing declaration fragment: ${fragment}`);
  }
});
