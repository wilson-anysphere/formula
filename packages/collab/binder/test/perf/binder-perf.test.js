/**
 * Perf regression benchmark: binder + (optional) session-style cell updates.
 *
 * Goal: Catch accidental O(N^2) regressions when the binder processes large batches
 * of cell changes.
 *
 * This is intentionally **opt-in** (skipped by default) so CI stays fast.
 *
 * How to run (recommended; uses the repo's TS-aware node:test harness):
 *   FORMULA_RUN_COLLAB_BINDER_PERF=1 \
 *     NODE_OPTIONS=--expose-gc \
 *     FORMULA_NODE_TEST_CONCURRENCY=1 \
 *     pnpm test:node binder-perf
 *
 * Direct (works on Node versions/configs that can execute workspace TypeScript imports):
 *   FORMULA_RUN_COLLAB_BINDER_PERF=1 \
 *     node --expose-gc --test --test-concurrency=1 \
 *     packages/collab/binder/test/perf/binder-perf.test.js
 *
 * Knobs:
 *   PERF_CELL_UPDATES=50000   # total updates to apply (default: 50000)
 *   PERF_BATCH_SIZE=1000      # updates per Yjs transaction (default: 1000)
 *   PERF_COLS=100             # controls row/col distribution (default: 100)
 *   PERF_KEY_ENCODING=canonical|legacy|rxc  # key format for Yjs writes (default: canonical)
 *   PERF_SCENARIO=yjs-to-dc|dc-to-yjs|all   # run only one scenario (default: all)
 *   PERF_INCLUDE_GUARDS=0     # set to 0 to disable canRead/canEdit hooks (default: enabled)
 *   PERF_TIMEOUT_MS=600000    # overall test timeout; also used for internal waits (default: 10 min)
 *
 * Optional CI-style enforcement (disabled unless set):
 *   PERF_MAX_TOTAL_MS_YJS_TO_DC=15000   # fail if total runtime exceeds this (ms)
 *   PERF_MAX_TOTAL_MS_DC_TO_YJS=15000   # fail if total runtime exceeds this (ms)
 *   PERF_MAX_PEAK_HEAP_BYTES_YJS_TO_DC=500000000 # fail if heapUsed peak exceeds this (bytes)
 *   PERF_MAX_PEAK_RSS_BYTES_YJS_TO_DC=1500000000 # fail if rss peak exceeds this (bytes)
 *   PERF_MAX_PEAK_HEAP_BYTES_DC_TO_YJS=500000000 # fail if heapUsed peak exceeds this (bytes)
 *   PERF_MAX_PEAK_RSS_BYTES_DC_TO_YJS=1500000000 # fail if rss peak exceeds this (bytes)
 *
 * Optional structured output:
 *   PERF_JSON=1  # emit JSON objects (one per scenario) for easy CI parsing
 */

import test from "node:test";
import assert from "node:assert/strict";
import { performance } from "node:perf_hooks";

const RUN_PERF = process.env.FORMULA_RUN_COLLAB_BINDER_PERF === "1";
const SCENARIO_FILTER = (process.env.PERF_SCENARIO ?? "").trim().toLowerCase();

/**
 * @param {"yjs-to-dc" | "dc-to-yjs"} scenario
 */
function shouldRunScenario(scenario) {
  if (!SCENARIO_FILTER || SCENARIO_FILTER === "all") return true;
  const normalized =
    SCENARIO_FILTER === "yjs->dc" || SCENARIO_FILTER === "yjs_to_dc"
      ? "yjs-to-dc"
      : SCENARIO_FILTER === "dc->yjs" || SCENARIO_FILTER === "dc_to_yjs"
        ? "dc-to-yjs"
        : SCENARIO_FILTER;
  return normalized === scenario;
}

/**
 * @param {"yjs-to-dc" | "dc-to-yjs"} scenario
 */
function perfTestForScenario(scenario) {
  if (!RUN_PERF) return test.skip;
  if (!shouldRunScenario(scenario)) return test.skip;
  return test;
}

const perfTestYjsToDc = perfTestForScenario("yjs-to-dc");
const perfTestDcToYjs = perfTestForScenario("dc-to-yjs");

const INCLUDE_GUARDS = process.env.PERF_INCLUDE_GUARDS !== "0";

function readPositiveInt(value, fallback) {
  const n = Number.parseInt(value ?? "", 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

const PERF_TIMEOUT_MS = readPositiveInt(process.env.PERF_TIMEOUT_MS, 10 * 60_000);

async function waitForCondition(fn, timeoutMs = PERF_TIMEOUT_MS) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      if (await fn()) return;
    } catch {
      // ignore while polling
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

function formatBytes(bytes) {
  const n = Number(bytes);
  if (!Number.isFinite(n)) return String(bytes);
  if (n < 1024) return `${n} B`;
  const kb = n / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KiB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MiB`;
  const gb = mb / 1024;
  return `${gb.toFixed(2)} GiB`;
}

function runtimeInfo() {
  return { node: process.version, platform: process.platform, arch: process.arch };
}

class StyleTableStub {
  constructor() {
    /** @type {Map<string, number>} */
    this._idsByKey = new Map();
    this._nextId = 1;
  }

  /**
   * @param {any} format
   */
  intern(format) {
    const key = JSON.stringify(format);
    const existing = this._idsByKey.get(key);
    if (existing) return existing;
    const id = this._nextId++;
    this._idsByKey.set(key, id);
    return id;
  }

  /**
   * @param {number} id
   */
  get(id) {
    return null;
  }
}

class DocumentControllerPerfStub {
  constructor() {
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this._listeners = new Map();
    /** @type {Map<string, { value: any, formula: string | null, styleId: number }>} */
    this._cells = new Map();

    this.styleTable = new StyleTableStub();
    // `bindYjsToDocumentController` patches this hook when edit guards are present.
    this.canEditCell = null;

    /** @type {number} */
    this.appliedDeltaCount = 0;
    /** @type {Array<{ target: number, resolve: () => void }>} */
    this._waiters = [];
  }

  /**
   * @param {string} event
   * @param {(payload: any) => void} listener
   */
  on(event, listener) {
    let set = this._listeners.get(event);
    if (!set) {
      set = new Set();
      this._listeners.set(event, set);
    }
    set.add(listener);
    return () => set.delete(listener);
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  _emit(event, payload) {
    const set = this._listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }

  /**
   * @param {string} sheetId
   * @param {{ row: number, col: number }} coord
   */
  getCell(sheetId, coord) {
    const key = `${sheetId}:${coord.row}:${coord.col}`;
    return this._cells.get(key) ?? { value: null, formula: null, styleId: 0 };
  }

  /**
   * Binder may call this during sheet metadata hydration when used via session wiring.
   * Keep it cheap and deterministic.
   */
  getSheetView(_sheetId) {
    return { frozenRows: 0, frozenCols: 0 };
  }

  /**
   * @param {any[]} deltas
   */
  applyExternalDeltas(deltas) {
    if (!Array.isArray(deltas) || deltas.length === 0) return;
    for (const delta of deltas) {
      const key = `${delta.sheetId}:${delta.row}:${delta.col}`;
      this._cells.set(key, {
        value: delta.after?.value ?? null,
        formula: delta.after?.formula ?? null,
        styleId: Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0,
      });
    }
    this.appliedDeltaCount += deltas.length;
    this._notifyWaiters();
  }

  /**
   * @param {number} target
   */
  whenApplied(target) {
    if (this.appliedDeltaCount >= target) return Promise.resolve();
    return new Promise((resolve) => {
      this._waiters.push({ target, resolve });
    });
  }

  _notifyWaiters() {
    if (this._waiters.length === 0) return;
    const remaining = [];
    for (const waiter of this._waiters) {
      if (this.appliedDeltaCount >= waiter.target) waiter.resolve();
      else remaining.push(waiter);
    }
    this._waiters = remaining;
  }
}

perfTestYjsToDc(
  "perf: binder applies many cell updates without pathological scaling",
  { timeout: PERF_TIMEOUT_MS, concurrency: 1 },
  async () => {
    const totalUpdates = Number.parseInt(process.env.PERF_CELL_UPDATES ?? "50000", 10);
    const batchSize = Number.parseInt(process.env.PERF_BATCH_SIZE ?? "1000", 10);
    const cols = Number.parseInt(process.env.PERF_COLS ?? "100", 10);
    const keyEncoding = (process.env.PERF_KEY_ENCODING ?? "canonical").trim();

    if (!Number.isFinite(totalUpdates) || totalUpdates <= 0) throw new Error("PERF_CELL_UPDATES must be a positive integer");
    if (!Number.isFinite(batchSize) || batchSize <= 0) throw new Error("PERF_BATCH_SIZE must be a positive integer");
    if (!Number.isFinite(cols) || cols <= 0) throw new Error("PERF_COLS must be a positive integer");
    if (!["canonical", "legacy", "rxc"].includes(keyEncoding)) {
      throw new Error('PERF_KEY_ENCODING must be one of: "canonical", "legacy", "rxc"');
    }

    const [{ bindYjsToDocumentController }, Y] = await Promise.all([import("../../index.js"), import("yjs")]);

    const ydoc = new Y.Doc();
    const dc = new DocumentControllerPerfStub();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController: dc,
      defaultSheetId: "Sheet1",
      ...(INCLUDE_GUARDS
        ? {
            // CollabSession wiring always provides these hooks; include them by default so
            // we exercise the guarded binder paths (permission/encryption checks).
            canReadCell: () => true,
            canEditCell: () => true,
          }
        : {}),
    });

    try {
      const cells = ydoc.getMap("cells");
      const origin = { type: "perf-origin" };

      if (RUN_PERF && typeof global.gc !== "function") {
        console.warn(
          "[binder-perf] global.gc() unavailable; run with NODE_OPTIONS=--expose-gc for more stable memory readings",
        );
      }

      if (typeof global.gc === "function") global.gc();
      const startMem = process.memoryUsage();
      let peakHeapUsed = startMem.heapUsed;
      let peakRss = startMem.rss;

      const t0 = performance.now();
      for (let i = 0; i < totalUpdates; i += batchSize) {
        const end = Math.min(totalUpdates, i + batchSize);

        ydoc.transact(() => {
          for (let j = i; j < end; j += 1) {
            const row = Math.floor(j / cols);
            const col = j % cols;
            const key =
              keyEncoding === "legacy"
                ? `Sheet1:${row},${col}`
                : keyEncoding === "rxc"
                  ? `r${row}c${col}`
                  : `Sheet1:${row}:${col}`;

            let cell = cells.get(key);
            // Handle the (rare) case where a prior write stored a non-Y.Map value.
            if (!(cell instanceof Y.Map)) {
              cell = new Y.Map();
              cells.set(key, cell);
            }

            // Approximate CollabSession's plain cell schema (value + explicit formula marker).
            cell.set("value", j);
            cell.set("formula", null);
          }
        }, origin);

        const mem = process.memoryUsage();
        peakHeapUsed = Math.max(peakHeapUsed, mem.heapUsed);
        peakRss = Math.max(peakRss, mem.rss);
      }
      const tWriteDone = performance.now();

      await dc.whenApplied(totalUpdates);
      const tApplyDone = performance.now();

      const endMem = process.memoryUsage();
      peakHeapUsed = Math.max(peakHeapUsed, endMem.heapUsed);
      peakRss = Math.max(peakRss, endMem.rss);

      // Best-effort: stabilize the final memory reading (requires `--expose-gc`).
      if (typeof global.gc === "function") global.gc();
      const postGcMem = process.memoryUsage();

      assert.equal(dc.appliedDeltaCount, totalUpdates);

      const writeMs = tWriteDone - t0;
      const applyMs = tApplyDone - tWriteDone;
      const totalMs = tApplyDone - t0;

      console.log(
        [
          "",
          `[binder-perf] updates=${totalUpdates.toLocaleString()} batchSize=${batchSize.toLocaleString()} cols=${cols.toLocaleString()} keyEncoding=${keyEncoding}`,
          `[binder-perf] time: write=${writeMs.toFixed(1)}ms apply=${applyMs.toFixed(1)}ms total=${totalMs.toFixed(1)}ms`,
          `[binder-perf] mem (best-effort): heapUsed start=${formatBytes(startMem.heapUsed)} peak=${formatBytes(peakHeapUsed)} postGC=${formatBytes(postGcMem.heapUsed)}`,
          `[binder-perf] mem (best-effort): rss      start=${formatBytes(startMem.rss)} peak=${formatBytes(peakRss)} postGC=${formatBytes(postGcMem.rss)}`,
        ].join("\n"),
      );

      if (process.env.PERF_JSON === "1") {
        console.log(
          JSON.stringify({
            suite: "binder-perf",
            scenario: "yjs->dc",
            runtime: runtimeInfo(),
            updates: totalUpdates,
            batchSize,
            cols,
            keyEncoding,
            timingMs: { write: writeMs, apply: applyMs, total: totalMs },
            mem: {
              heapUsed: { start: startMem.heapUsed, peak: peakHeapUsed, postGc: postGcMem.heapUsed },
              rss: { start: startMem.rss, peak: peakRss, postGc: postGcMem.rss },
            },
          }),
        );
      }

      const maxTotalMs = readPositiveInt(process.env.PERF_MAX_TOTAL_MS_YJS_TO_DC, 0);
      if (maxTotalMs > 0) {
        assert.ok(
          totalMs <= maxTotalMs,
          `[binder-perf] expected total <= ${maxTotalMs}ms, got ${totalMs.toFixed(1)}ms`,
        );
      }

      const maxPeakHeap = readPositiveInt(process.env.PERF_MAX_PEAK_HEAP_BYTES_YJS_TO_DC, 0);
      if (maxPeakHeap > 0) {
        assert.ok(
          peakHeapUsed <= maxPeakHeap,
          `[binder-perf] expected peak heapUsed <= ${maxPeakHeap} bytes, got ${peakHeapUsed}`,
        );
      }
      const maxPeakRss = readPositiveInt(process.env.PERF_MAX_PEAK_RSS_BYTES_YJS_TO_DC, 0);
      if (maxPeakRss > 0) {
        assert.ok(peakRss <= maxPeakRss, `[binder-perf] expected peak rss <= ${maxPeakRss} bytes, got ${peakRss}`);
      }
    } finally {
      try {
        binder.destroy();
      } catch {
        // ignore
      }
      try {
        ydoc.destroy();
      } catch {
        // ignore
      }
    }
  },
);

perfTestDcToYjs(
  "perf: binder writes many DocumentController deltas to Yjs without pathological scaling",
  { timeout: PERF_TIMEOUT_MS, concurrency: 1 },
  async () => {
    const totalUpdates = Number.parseInt(process.env.PERF_CELL_UPDATES ?? "50000", 10);
    const batchSize = Number.parseInt(process.env.PERF_BATCH_SIZE ?? "1000", 10);
    const cols = Number.parseInt(process.env.PERF_COLS ?? "100", 10);

    if (!Number.isFinite(totalUpdates) || totalUpdates <= 0) throw new Error("PERF_CELL_UPDATES must be a positive integer");
    if (!Number.isFinite(batchSize) || batchSize <= 0) throw new Error("PERF_BATCH_SIZE must be a positive integer");
    if (!Number.isFinite(cols) || cols <= 0) throw new Error("PERF_COLS must be a positive integer");

    const [{ bindYjsToDocumentController }, Y] = await Promise.all([import("../../index.js"), import("yjs")]);

    const ydoc = new Y.Doc();
    const dc = new DocumentControllerPerfStub();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController: dc,
      defaultSheetId: "Sheet1",
      ...(INCLUDE_GUARDS
        ? {
            canReadCell: () => true,
            canEditCell: () => true,
          }
        : {}),
    });

    try {
      const cells = ydoc.getMap("cells");

      if (RUN_PERF && typeof global.gc !== "function") {
        console.warn(
          "[binder-perf] global.gc() unavailable; run with NODE_OPTIONS=--expose-gc for more stable memory readings",
        );
      }

      if (typeof global.gc === "function") global.gc();
      const startMem = process.memoryUsage();
      let peakHeapUsed = startMem.heapUsed;
      let peakRss = startMem.rss;

      const t0 = performance.now();
      for (let i = 0; i < totalUpdates; i += batchSize) {
        const end = Math.min(totalUpdates, i + batchSize);

        /** @type {any[]} */
        const deltas = [];
        for (let j = i; j < end; j += 1) {
          const row = Math.floor(j / cols);
          const col = j % cols;
          deltas.push({
            sheetId: "Sheet1",
            row,
            col,
            before: { value: null, formula: null, styleId: 0 },
            after: { value: j, formula: null, styleId: 0 },
          });
        }

        dc._emit("change", { deltas });

        const mem = process.memoryUsage();
        peakHeapUsed = Math.max(peakHeapUsed, mem.heapUsed);
        peakRss = Math.max(peakRss, mem.rss);
      }
      const tEmitDone = performance.now();

      await waitForCondition(() => cells.size === totalUpdates);
      const tWriteDone = performance.now();

      const endMem = process.memoryUsage();
      peakHeapUsed = Math.max(peakHeapUsed, endMem.heapUsed);
      peakRss = Math.max(peakRss, endMem.rss);

      if (typeof global.gc === "function") global.gc();
      const postGcMem = process.memoryUsage();

      assert.equal(cells.size, totalUpdates);

      const emitMs = tEmitDone - t0;
      const writeMs = tWriteDone - tEmitDone;
      const totalMs = tWriteDone - t0;

      console.log(
        [
          "",
          `[binder-perf] (doc->yjs) updates=${totalUpdates.toLocaleString()} batchSize=${batchSize.toLocaleString()} cols=${cols.toLocaleString()}`,
          `[binder-perf] time: emit=${emitMs.toFixed(1)}ms binderWrite=${writeMs.toFixed(1)}ms total=${totalMs.toFixed(1)}ms`,
          `[binder-perf] mem (best-effort): heapUsed start=${formatBytes(startMem.heapUsed)} peak=${formatBytes(peakHeapUsed)} postGC=${formatBytes(postGcMem.heapUsed)}`,
          `[binder-perf] mem (best-effort): rss      start=${formatBytes(startMem.rss)} peak=${formatBytes(peakRss)} postGC=${formatBytes(postGcMem.rss)}`,
        ].join("\n"),
      );

      if (process.env.PERF_JSON === "1") {
        console.log(
          JSON.stringify({
            suite: "binder-perf",
            scenario: "dc->yjs",
            runtime: runtimeInfo(),
            updates: totalUpdates,
            batchSize,
            cols,
            timingMs: { emit: emitMs, binderWrite: writeMs, total: totalMs },
            mem: {
              heapUsed: { start: startMem.heapUsed, peak: peakHeapUsed, postGc: postGcMem.heapUsed },
              rss: { start: startMem.rss, peak: peakRss, postGc: postGcMem.rss },
            },
          }),
        );
      }

      const maxTotalMs = readPositiveInt(process.env.PERF_MAX_TOTAL_MS_DC_TO_YJS, 0);
      if (maxTotalMs > 0) {
        assert.ok(
          totalMs <= maxTotalMs,
          `[binder-perf] expected total <= ${maxTotalMs}ms, got ${totalMs.toFixed(1)}ms`,
        );
      }

      const maxPeakHeap = readPositiveInt(process.env.PERF_MAX_PEAK_HEAP_BYTES_DC_TO_YJS, 0);
      if (maxPeakHeap > 0) {
        assert.ok(
          peakHeapUsed <= maxPeakHeap,
          `[binder-perf] expected peak heapUsed <= ${maxPeakHeap} bytes, got ${peakHeapUsed}`,
        );
      }
      const maxPeakRss = readPositiveInt(process.env.PERF_MAX_PEAK_RSS_BYTES_DC_TO_YJS, 0);
      if (maxPeakRss > 0) {
        assert.ok(peakRss <= maxPeakRss, `[binder-perf] expected peak rss <= ${maxPeakRss} bytes, got ${peakRss}`);
      }
    } finally {
      try {
        binder.destroy();
      } catch {
        // ignore
      }
      try {
        ydoc.destroy();
      } catch {
        // ignore
      }
    }
  },
);
