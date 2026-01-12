import assert from "node:assert/strict";
import test from "node:test";

import { createSheetNameResolverFromIdToNameMap } from "./sheetNameResolver.ts";

test("SheetNameResolver resolves display names with Unicode NFKC + case-insensitive compare", () => {
  const sheetIdToName = new Map([
    ["sheet-1", "Budget"],
    ["sheet-2", "é"],
    // Angstrom sign (U+212B) normalizes to Å (U+00C5) under NFKC.
    ["sheet-3", "Å"],
  ]);

  const resolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

  assert.equal(resolver.getSheetIdByName("É"), "sheet-2");
  assert.equal(resolver.getSheetIdByName("Å"), "sheet-3");
});

test("SheetNameResolver accepts ids directly (case-insensitive)", () => {
  const sheetIdToName = new Map([["Sheet2", "Data"]]);
  const resolver = createSheetNameResolverFromIdToNameMap(sheetIdToName);

  assert.equal(resolver.getSheetIdByName("sheet2"), "Sheet2");
  assert.equal(resolver.getSheetNameById("sheet2"), "Data");
  assert.equal(resolver.getSheetIdByName("Data"), "Sheet2");
});

