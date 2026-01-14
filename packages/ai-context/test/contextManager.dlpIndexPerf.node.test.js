import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

function makePolicy({ maxAllowed = "Internal", redactDisallowed = true } = {}) {
  return {
    version: 1,
    allowDocumentOverrides: true,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed,
        allowRestrictedContent: false,
        redactDisallowed,
      },
    },
  };
}

function instrumentIterationPasses(records) {
  let passes = 0;
  const proxy = new Proxy(records, {
    get(target, prop, receiver) {
      if (prop === Symbol.iterator) {
        return function () {
          passes += 1;
          return target[Symbol.iterator]();
        };
      }
      return Reflect.get(target, prop, receiver);
    },
  });
  return { proxy, getPasses: () => passes };
}

test("ContextManager.buildContext: avoids scanning classification records per cell under REDACT decisions", async () => {
  const ragIndex = {
    // Minimal store shape; ContextManager only reads `.size` in a few tests/logs.
    store: { size: 0 },
    async indexSheet(sheet) {
      return {
        schema: { name: sheet.name, tables: [], namedRanges: [], dataRegions: [] },
        chunkCount: 0,
      };
    },
    async search() {
      return [];
    },
  };

  const cm = new ContextManager({
    tokenBudgetTokens: 1_000,
    // Keep outputs stable; not relevant to this perf proxy.
    redactor: (t) => t,
    ragIndex: /** @type {any} */ (ragIndex),
    cacheSheetIndex: false,
  });

  // Use a moderately sized sheet window so a per-cell `effectiveCellClassification` regression
  // would cause many passes over the record list.
  const rows = 50;
  const cols = 50;
  const values = Array.from({ length: rows }, () => Array.from({ length: cols }, () => null));

  const documentId = "doc-dlp-index";
  const sheetId = "Sheet1";

  // A single Confidential cell selector is enough to trigger a structured REDACT decision when
  // `maxAllowed=Internal`. If a regression falls back to per-cell scanning, we'd see one full pass
  // over the record list per cell (~10k passes).
  const records = [
    {
      selector: { scope: "cell", documentId, sheetId, row: 0, col: 0 },
      classification: { level: "Confidential", labels: [] },
    },
  ];
  const { proxy: recordsProxy, getPasses } = instrumentIterationPasses(records);

  const out = await cm.buildContext({
    sheet: { name: sheetId, values },
    query: "anything",
    sampleRows: 1,
    samplingStrategy: "head",
    dlp: {
      documentId,
      sheetId,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: recordsProxy,
    },
  });

  assert.equal(out.sampledRows[0][0], "[REDACTED]");

  // Expect only a handful of linear passes (sheet/doc scan + range scan + index build).
  // Any per-cell scanning regression would exceed this by orders of magnitude.
  assert.ok(getPasses() < 50, `expected < 50 record iteration passes, got ${getPasses()}`);
});
