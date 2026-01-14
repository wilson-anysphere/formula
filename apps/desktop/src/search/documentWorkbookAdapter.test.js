import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../document/documentController.js";
import { DocumentWorkbookAdapter } from "./documentWorkbookAdapter.js";

test("DocumentWorkbookAdapter preserves defined-name case while remaining case-insensitive", () => {
  const workbook = new DocumentWorkbookAdapter({ document: { getSheetIds: () => [] } });
  const range = { startRow: 0, endRow: 9, startCol: 0, endCol: 0 };

  workbook.defineName("SalesData", { sheetName: "Sheet1", range });

  const storedLower = workbook.getName("salesdata");
  assert.ok(storedLower);
  assert.equal(storedLower.name, "SalesData");
  assert.deepEqual(storedLower.range, range);

  const storedUpper = workbook.getName("SALESDATA");
  assert.ok(storedUpper);
  assert.equal(storedUpper.name, "SalesData");
  assert.deepEqual(storedUpper.range, range);
});

test("DocumentWorkbookAdapter schemaVersion bumps when schema changes", () => {
  const workbook = new DocumentWorkbookAdapter({ document: { getSheetIds: () => [] } });
  assert.equal(workbook.schemaVersion, 0);

  workbook.defineName("Name1", { sheetName: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
  assert.equal(workbook.schemaVersion, 1);

  workbook.addTable({ name: "Table1", sheetName: "Sheet1", startRow: 0, endRow: 1, startCol: 0, endCol: 1, columns: ["A"] });
  assert.equal(workbook.schemaVersion, 2);

  workbook.clearSchema();
  assert.equal(workbook.schemaVersion, 3);
  assert.equal(workbook.names.size, 0);
  assert.equal(workbook.tables.size, 0);
});

test("DocumentWorkbookAdapter resolves sheets by display name via sheetNameResolver", () => {
  const doc = { getSheetIds: () => ["Sheet1", "Sheet2"] };
  const namesById = new Map([
    ["Sheet1", "Sheet1"],
    ["Sheet2", "My Sheet"],
  ]);
  const sheetNameResolver = {
    getSheetNameById: (id) => namesById.get(id) ?? null,
    getSheetIdByName: (name) => {
      const needle = String(name ?? "").trim().toLowerCase();
      if (!needle) return null;
      for (const [id, sheetName] of namesById.entries()) {
        if (sheetName.toLowerCase() === needle) return id;
      }
      return null;
    },
  };

  const workbook = new DocumentWorkbookAdapter({ document: doc, sheetNameResolver });

  const sheetNames = workbook.sheets.map((s) => s.name);
  assert.deepEqual(sheetNames, ["Sheet1", "My Sheet"]);

  const sheet = workbook.getSheet("My Sheet");
  assert.equal(sheet.sheetId, "Sheet2");

  const sheetLower = workbook.getSheet("my sheet");
  assert.equal(sheetLower.sheetId, "Sheet2");

  // Accept Excel-style quoting for sheet tokens (e.g. from sheet-qualified references).
  const sheetQuoted = workbook.getSheet("'My Sheet'");
  assert.equal(sheetQuoted.sheetId, "Sheet2");

  assert.throws(() => workbook.getSheet("MissingSheet"), /Unknown sheet/i);
});

test("DocumentWorkbookAdapter uses stable sheetId for cell access when sheets are renamed", () => {
  const doc = new DocumentController();
  // Stable id does not match display name.
  doc.setCellValue("sheet-1", "A1", "hello");

  const namesById = new Map([["sheet-1", "Budget"]]);
  const workbook = new DocumentWorkbookAdapter({
    document: doc,
    sheetNameResolver: {
      getSheetNameById: (id) => namesById.get(String(id)) ?? null,
      getSheetIdByName: (name) => {
        const needle = String(name ?? "").trim().toLowerCase();
        if (!needle) return null;
        for (const [id, sheetName] of namesById.entries()) {
          if (sheetName.toLowerCase() === needle) return id;
        }
        return null;
      },
    },
  });

  const sheet = workbook.getSheet("Budget");
  assert.equal(sheet.sheetId, "sheet-1");
  // Ensure we read from the stable id (and do not create a phantom "Budget" sheet).
  const cell = sheet.getCell(0, 0);
  assert.ok(cell);
  assert.equal(cell.value, "hello");
  assert.deepEqual(doc.getSheetIds(), ["sheet-1"]);
});

test("DocumentWorkbookAdapter trims sheet ids returned by sheetNameResolver", () => {
  const doc = new DocumentController();
  doc.setCellValue("sheet-1", "A1", "hello");

  const sheetNameResolver = {
    getSheetNameById: (id) => (String(id) === "sheet-1" ? "Budget" : null),
    // Some integration layers may accidentally include whitespace around ids.
    getSheetIdByName: (name) => (String(name ?? "").trim().toLowerCase() === "budget" ? "  sheet-1  " : null),
  };

  const workbook = new DocumentWorkbookAdapter({ document: doc, sheetNameResolver });
  const sheet = workbook.getSheet("Budget");
  assert.equal(sheet.sheetId, "sheet-1");

  const cell = sheet.getCell(0, 0);
  assert.ok(cell);
  assert.equal(cell.value, "hello");
  // Ensure we did not create a phantom sheet keyed by the untrimmed id.
  assert.deepEqual(doc.getSheetIds(), ["sheet-1"]);
});

test("DocumentWorkbookAdapter does not create phantom sheets when name resolution fails", () => {
  const doc = new DocumentController();
  doc.setCellValue("sheet-1", "A1", "hello");

  const namesById = new Map([["sheet-1", "Budget"]]);
  const workbook = new DocumentWorkbookAdapter({
    document: doc,
    sheetNameResolver: {
      getSheetNameById: (id) => namesById.get(String(id)) ?? null,
      getSheetIdByName: (name) => {
        const needle = String(name ?? "").trim().toLowerCase();
        if (!needle) return null;
        for (const [id, sheetName] of namesById.entries()) {
          if (sheetName.toLowerCase() === needle) return id;
        }
        return null;
      },
    },
  });

  assert.throws(() => workbook.getSheet("DoesNotExist"), /Unknown sheet/i);
  assert.deepEqual(doc.getSheetIds(), ["sheet-1"]);
});

test("DocumentWorkbookAdapter resolves quoted sheet names via sheetNameResolver", () => {
  const doc = { getSheetIds: () => ["Sheet1", "Sheet2"] };
  const namesById = new Map([
    ["Sheet1", "Budget"],
    ["Sheet2", "O'Brien"],
  ]);
  const sheetNameResolver = {
    getSheetNameById: (id) => namesById.get(id) ?? null,
    getSheetIdByName: (name) => {
      const needle = String(name ?? "").trim().toLowerCase();
      if (!needle) return null;
      for (const [id, sheetName] of namesById.entries()) {
        if (sheetName.toLowerCase() === needle) return id;
      }
      return null;
    },
  };

  const workbook = new DocumentWorkbookAdapter({ document: doc, sheetNameResolver });

  assert.equal(workbook.getSheet("'Budget'").sheetId, "Sheet1");
  assert.equal(workbook.getSheet("'O''Brien'").sheetId, "Sheet2");
  assert.throws(() => workbook.getSheet("'Missing'"), /Unknown sheet/i);
});

test("DocumentWorkbookAdapter does not create phantom sheets after a rename (stale display name)", () => {
  const doc = new DocumentController();
  doc.setCellValue("sheet-1", "A1", "hello");

  // Mutable mapping to simulate rename.
  const namesById = new Map([["sheet-1", "Budget"]]);
  const sheetNameResolver = {
    getSheetNameById: (id) => namesById.get(String(id)) ?? null,
    getSheetIdByName: (name) => {
      const needle = String(name ?? "").trim().toLowerCase();
      if (!needle) return null;
      for (const [id, sheetName] of namesById.entries()) {
        if (sheetName.toLowerCase() === needle) return id;
      }
      return null;
    },
  };

  const workbook = new DocumentWorkbookAdapter({ document: doc, sheetNameResolver });

  // Initially resolvable.
  assert.equal(workbook.getSheet("Budget").sheetId, "sheet-1");

  // Rename display name (id stays stable).
  namesById.set("sheet-1", "Budget2026");

  // Old display name should no longer resolve and must not create a new sheet.
  assert.throws(() => workbook.getSheet("Budget"), /Unknown sheet/i);
  assert.deepEqual(doc.getSheetIds(), ["sheet-1"]);

  // New display name should resolve to the same sheet id.
  const renamed = workbook.getSheet("Budget2026");
  assert.equal(renamed.sheetId, "sheet-1");
  assert.equal(renamed.name, "Budget2026");

  // Workbook sheet listing should also reflect the updated display name.
  assert.deepEqual(workbook.sheets.map((s) => s.name), ["Budget2026"]);
});

test("DocumentWorkbookAdapter resolves display names for unmaterialized sheets without creating them", () => {
  const doc = new DocumentController();

  // DocumentController starts with no sheets until they're referenced by cell operations
  // or added via sheet metadata operations.
  assert.deepEqual(doc.getSheetIds(), []);

  // Simulate a sheet that exists in external metadata (e.g. sheet tabs) but hasn't been
  // materialized in the DocumentController yet.
  const namesById = new Map([["sheet-1", "Budget"]]);
  const workbook = new DocumentWorkbookAdapter({
    document: doc,
    sheetNameResolver: {
      getSheetNameById: (id) => namesById.get(String(id)) ?? null,
      getSheetIdByName: (name) => {
        const needle = String(name ?? "").trim().toLowerCase();
        if (!needle) return null;
        for (const [id, sheetName] of namesById.entries()) {
          if (sheetName.toLowerCase() === needle) return id;
        }
        return null;
      },
    },
  });

  const sheet = workbook.getSheet("Budget");
  assert.equal(sheet.sheetId, "sheet-1");

  // `getSheet` should not create a sheet in the DocumentController; it only resolves names.
  assert.deepEqual(doc.getSheetIds(), []);

  // Accessors that probe for cell values should also avoid materializing the sheet.
  assert.equal(sheet.getCell(0, 0), null);
  assert.deepEqual(doc.getSheetIds(), []);

  // Accessors that are used by search (usedRange iteration) should also avoid materializing.
  assert.equal(sheet.getUsedRange(), null);
  assert.deepEqual(doc.getSheetIds(), []);
});

test("DocumentWorkbookAdapter resolves DocumentController sheet meta names with Unicode NFKC + case-insensitive compare", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.renameSheet("Sheet1", "Å");

  // No explicit `sheetNameResolver` passed, so the adapter falls back to DocumentController metadata.
  const workbook = new DocumentWorkbookAdapter({ document: doc });

  // Angstrom sign (U+212B) normalizes to Å (U+00C5) under NFKC.
  const sheet = workbook.getSheet("Å");
  assert.equal(sheet.sheetId, "Sheet1");
  assert.equal(sheet.name, "Å");
});

test("DocumentWorkbookAdapter exposes merged-cell metadata for search semantics", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "hello");
  doc.setMergedRanges("Sheet1", [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }], { label: "Merge Cells" }); // A1:B1

  const workbook = new DocumentWorkbookAdapter({ document: doc });
  const sheet = workbook.getSheet("Sheet1");

  assert.deepEqual(sheet.getMergedRanges(), [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }]);
  assert.deepEqual(sheet.getMergedMasterCell(0, 1), { row: 0, col: 0 });
  assert.equal(sheet.getMergedMasterCell(5, 5), null);
});

test("DocumentWorkbookAdapter formats rich text + image payloads for search display", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { text: "Hello", runs: [{ start: 0, end: 5, style: { bold: true } }] });
  doc.setCellValue("Sheet1", "A2", { type: "image", value: { imageId: "img-1", altText: "Logo" } });
  doc.setCellValue("Sheet1", "A3", { type: "image", value: { imageId: "img-2" } });

  const workbook = new DocumentWorkbookAdapter({ document: doc });
  const sheet = workbook.getSheet("Sheet1");

  assert.equal(sheet.getCell(0, 0).display, "Hello");
  assert.equal(sheet.getCell(1, 0).display, "Logo");
  assert.equal(sheet.getCell(2, 0).display, "[Image]");
});

test("DocumentWorkbookAdapter iterateCells exposes rich text + image display strings", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", { type: "image", value: { imageId: "img-1", altText: "Logo" } });
  doc.setCellValue("Sheet1", "B2", { text: "Bold", runs: [{ start: 0, end: 4, style: { bold: true } }] });

  const workbook = new DocumentWorkbookAdapter({ document: doc });
  const sheet = workbook.getSheet("Sheet1");

  const entries = Array.from(sheet.iterateCells({ startRow: 0, endRow: 10, startCol: 0, endCol: 10 }));
  const byCoord = new Map(entries.map((e) => [`${e.row},${e.col}`, e.cell.display]));
  assert.equal(byCoord.get("0,0"), "Logo");
  assert.equal(byCoord.get("1,1"), "Bold");
});
