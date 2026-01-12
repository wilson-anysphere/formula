import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { MacroRecorder } from "../apps/desktop/src/macro-recorder/index.js";
import { DocumentControllerWorkbookAdapter } from "../apps/desktop/src/scripting/documentControllerWorkbookAdapter.js";

test("Macro recorder handles DocumentController formatDeltas (layered column formatting)", () => {
  const doc = new DocumentController();
  const workbook = new DocumentControllerWorkbookAdapter(doc, { activeSheetName: "Sheet1" });

  const recorder = new MacroRecorder(workbook);
  recorder.start();

  doc.setColFormat("Sheet1", 2, { font: { bold: true } });

  const actions = recorder.stop();
  assert.deepEqual(actions, [{ type: "setFormat", sheetName: "Sheet1", address: "C1:C1048576", format: { bold: true } }]);

  workbook.dispose();
});

