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

function instrumentRecordList(records) {
  let passes = 0;
  let elementGets = 0;
  let propGets = 0;
  /** @type {WeakMap<object, any>} */
  const objectProxyCache = new WeakMap();

  /**
   * Wrap a plain object in a Proxy that counts property reads (recursively).
   *
   * This catches regressions where callers clone the record *array* once (e.g. `Array.from(...)`)
   * and then scan the cloned array per-cell/per-hit, which would bypass array-level iteration
   * counters.
   *
   * @param {any} value
   * @returns {any}
   */
  function wrapObject(value) {
    if (!value || typeof value !== "object") return value;
    if (Array.isArray(value)) return value;
    const cached = objectProxyCache.get(value);
    if (cached) return cached;
    const proxy = new Proxy(value, {
      get(target, prop, receiver) {
        propGets += 1;
        return wrapObject(Reflect.get(target, prop, receiver));
      },
    });
    objectProxyCache.set(value, proxy);
    return proxy;
  }

  const wrappedRecords = (records ?? []).map((r) => wrapObject(r));
  const proxy = new Proxy(wrappedRecords, {
    get(target, prop, receiver) {
      if (prop === Symbol.iterator) {
        return function () {
          passes += 1;
          // Bind the iterator to the proxy so element reads go through this trap.
          return Array.prototype[Symbol.iterator].call(receiver);
        };
      }
      if (typeof prop === "string" && /^[0-9]+$/.test(prop)) {
        elementGets += 1;
      }
      return Reflect.get(target, prop, receiver);
    },
  });
  return { proxy, getPasses: () => passes, getElementGets: () => elementGets, getPropGets: () => propGets };
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
  const { proxy: recordsProxy, getPasses, getElementGets, getPropGets } = instrumentRecordList(records);

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
  // Defense-in-depth: catch per-cell scans even if implemented without `for..of` / Symbol.iterator.
  assert.ok(getElementGets() < 200, `expected < 200 record element reads, got ${getElementGets()}`);
  // Defense-in-depth: catch per-cell scans even if the record list is cloned once and scanned repeatedly.
  assert.ok(getPropGets() < 500, `expected < 500 record property reads, got ${getPropGets()}`);
});
