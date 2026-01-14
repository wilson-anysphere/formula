import assert from "node:assert/strict";
import test from "node:test";

import { AiContextManager } from "../src/aiContextManager.js";
import { DLP_ACTION } from "../src/actions.js";

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
   * This catches regressions where the record list is cloned once and then scanned
   * per cell, bypassing array-level scan counters.
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
          // Bind iterator to proxy so numeric index access is observable.
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

test("AiContextManager.buildCloudContext: avoids scanning classification records per cell under REDACT decisions", () => {
  const documentId = "doc-dlp-index";
  const sheetId = "Sheet1";

  const records = [
    {
      selector: { scope: "cell", documentId, sheetId, row: 0, col: 0 },
      classification: { level: "Restricted", labels: [] },
    },
  ];
  const { proxy: recordsProxy, getPasses, getElementGets, getPropGets } = instrumentRecordList(records);

  const auditEvents = [];
  const classificationStore = { list: () => recordsProxy };
  const auditLogger = { log: (event) => auditEvents.push(event) };
  const manager = new AiContextManager({ classificationStore, auditLogger });

  const out = manager.buildCloudContext({
    documentId,
    sheetId,
    range: { start: { row: 0, col: 0 }, end: { row: 49, col: 49 } },
    cells: [
      { row: 0, col: 0, value: "secret" },
      { row: 0, col: 1, value: "ok" },
    ],
    policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
  });

  assert.ok(out.context.includes("[REDACTED]"));
  assert.ok(out.context.includes("ok"));

  // Expect a small number of linear scans (selection classification + index build).
  // A per-cell scan regression would drive this into the thousands.
  assert.ok(getPasses() < 50, `expected < 50 record iteration passes, got ${getPasses()}`);
  assert.ok(getElementGets() < 200, `expected < 200 record element reads, got ${getElementGets()}`);
  assert.ok(getPropGets() < 500, `expected < 500 record property reads, got ${getPropGets()}`);

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.type, "ai.request");
  assert.equal(auditEvents[0]?.documentId, documentId);
  assert.equal(auditEvents[0]?.sheetId, sheetId);
});
