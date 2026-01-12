import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../../document/documentController.js";
import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.ts";
import { WorkbookContextBuilder } from "../WorkbookContextBuilder.ts";
import { createSheetNameResolverFromIdToNameMap } from "../../../sheet/sheetNameResolver.ts";

test("WorkbookContextBuilder formats sheet-qualified A1 refs using sheet display names (Excel quoting)", async () => {
  const documentController = new DocumentController();
  const sheetId = documentController.addSheet({ name: "O'Brien" });
  documentController.setRangeValues(sheetId, "A1", [["Name"], ["Alice"]]);

  const sheetIdToName = new Map([[sheetId, "O'Brien"]]);
  const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

  const spreadsheet = new DocumentControllerSpreadsheetApi(documentController, { sheetNameResolver });
  const builder = new WorkbookContextBuilder({
    workbookId: "wb_sheet_names",
    documentController,
    spreadsheet,
    ragService: null,
    sheetNameResolver,
    mode: "chat",
    model: "unit-test-model",
    maxSheets: 1,
    // Avoid trimming for deterministic assertions.
    maxPromptContextTokens: 1_000_000,
  });

  const ctx = await builder.build({ activeSheetId: sheetId });

  const sample = ctx.payload.blocks.find((b) => b.kind === "sheet_sample" && b.sheetId === sheetId);
  assert.ok(sample, "expected a sheet_sample block for the active sheet");
  assert.equal(sample.range, "'O''Brien'!A1:A2");

  const sheet = ctx.payload.sheets.find((s) => s.sheetId === sheetId);
  assert.ok(sheet, "expected a sheet summary for the active sheet");
  assert.equal(sheet.schema.name, "O'Brien");
  assert.equal(sheet.schema.dataRegions[0].range, "'O''Brien'!A1:A2");
});

test("WorkbookContextBuilder resolves retrieved sheet display names back to stable sheet ids", async () => {
  const documentController = new DocumentController();
  const mainId = documentController.addSheet({ name: "Main" });
  const dataId = documentController.addSheet({ name: "Data Sheet" });
  documentController.setRangeValues(mainId, "A1", [["Main"]]);
  documentController.setRangeValues(dataId, "A1", [["Data"]]);

  const sheetIdToName = new Map([
    [mainId, "Main"],
    [dataId, "Data Sheet"],
  ]);
  const sheetNameResolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

  const spreadsheet = new DocumentControllerSpreadsheetApi(documentController, { sheetNameResolver });
  const builder = new WorkbookContextBuilder({
    workbookId: "wb_sheet_names_retrieval",
    documentController,
    spreadsheet,
    ragService: {
      buildWorkbookContextFromSpreadsheetApi: async () => ({
        retrieved: [
          {
            id: "chunk_1",
            text: "Relevant content",
            metadata: { sheetName: "Data Sheet", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
          },
        ],
      }),
    },
    sheetNameResolver,
    mode: "chat",
    model: "unit-test-model",
    maxSheets: 2,
    maxRetrievedBlocks: 0,
    maxPromptContextTokens: 1_000_000,
  });

  const ctx = await builder.build({ activeSheetId: mainId, focusQuestion: "what is on the data sheet?" });
  const summarizedSheetIds = ctx.payload.sheets.map((s) => s.sheetId);
  assert.ok(summarizedSheetIds.includes(mainId), "expected active sheet id to be summarized");
  assert.ok(summarizedSheetIds.includes(dataId), "expected retrieved sheet to be summarized by stable id");
  assert.ok(!summarizedSheetIds.includes("Data Sheet"), "expected retrieved sheet name to be resolved to an id");

  assert.ok(ctx.payload.retrieval, "expected retrieval metadata to be present");
  assert.ok(
    ctx.payload.retrieval.retrievedRanges.includes("'Data Sheet'!A1"),
    "expected retrievedRanges to be formatted using the sheet display name",
  );
});

