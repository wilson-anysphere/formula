import { afterAll, describe, expect, it } from "vitest";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";

import {
  engineApplyDocumentChange,
  engineHydrateFromDocument,
  type EngineSyncTarget,
} from "../documentControllerSync.ts";
import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

function engineWorkerAsSyncTarget(engine: EngineWorker): EngineSyncTarget {
  // `engineHydrateFromDocument` / `engineApplyDocumentChange` operate on a sheet-first sync surface
  // (matching wasm-bindgen workbook signatures). `EngineWorker` exposes the public EngineClient API
  // which uses sheet-last for some calls. Adapt between the two for wasm integration tests.
  return {
    loadWorkbookFromJson: (json) => engine.loadWorkbookFromJson(json),
    setCell: (address, value, sheet) => engine.setCell(address, value, sheet),
    setCells: (updates) => engine.setCells(updates),
    recalculate: (sheet) => engine.recalculate(sheet),
    setSheetDisplayName: (sheetId, name) => engine.setSheetDisplayName(sheetId, name),
    setWorkbookFileMetadata: (directory, filename) => engine.setWorkbookFileMetadata(directory, filename),
  };
}

class MockWorkerGlobal {
  private readonly listeners = new Set<(event: MessageEvent<unknown>) => void>();

  addEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.add(listener);
  }

  dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    for (const listener of this.listeners) {
      listener(event);
    }
  }
}

class MockMessagePort {
  onmessage: ((event: MessageEvent<unknown>) => void) | null = null;
  private listeners = new Map<string, Set<(event: MessageEvent<unknown>) => void>>();
  private other: MockMessagePort | null = null;

  connect(other: MockMessagePort) {
    this.other = other;
  }

  postMessage(message: unknown): void {
    queueMicrotask(() => {
      this.other?.dispatchMessage(message);
    });
  }

  start(): void {}

  close(): void {
    this.listeners.clear();
    this.onmessage = null;
    this.other = null;
  }

  addEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    const key = String(type ?? "");
    let set = this.listeners.get(key);
    if (!set) {
      set = new Set();
      this.listeners.set(key, set);
    }
    set.add(listener);
  }

  removeEventListener(type: string, listener: (event: MessageEvent<unknown>) => void): void {
    const key = String(type ?? "");
    this.listeners.get(key)?.delete(listener);
  }

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners.get("message") ?? []) {
      listener(event);
    }
  }
}

function createMockChannel(): MessageChannelLike {
  const port1 = new MockMessagePort();
  const port2 = new MockMessagePort();
  port1.connect(port2);
  port2.connect(port1);
  return { port1, port2: port2 as unknown as MessagePort };
}

class LocalWorker implements WorkerLike {
  postMessage(message: unknown): void {
    workerGlobal.dispatchMessage(message);
  }

  terminate(): void {
    // No-op; the worker lives in-process for this test.
  }
}

describeWasm("DocumentController sheet renames → setSheetDisplayName → CELL() (wasm)", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it('uses sheet display names (tab names) for CELL("address") and CELL("filename") after renames', async () => {
    const wasmModuleUrl = new URL("./fixtures/formulaWasmNodeWrapper.mjs", import.meta.url).href;
    await loadWorkerModule();

    const engine = await EngineWorker.connect({
      worker: new LocalWorker(),
      wasmModuleUrl,
      channelFactory: createMockChannel,
    });

    try {
      const syncTarget = engineWorkerAsSyncTarget(engine);

      const doc = new DocumentController();
      doc.addSheet({ sheetId: "sheet_2", name: "Budget" });

      // Use stable sheet ids in formulas (`sheet_2!A1`) and rely on the engine's display-name metadata
      // for output formatting.
      doc.setCellFormula("Sheet1", "A1", 'CELL("address",sheet_2!A1)');
      doc.setCellFormula("Sheet1", "A2", 'CELL("filename",sheet_2!A1)');

      await engineHydrateFromDocument(syncTarget, doc, {
        workbookFileMetadata: { directory: null, filename: "book.xlsx" },
      });

      let cell = await engine.getCell("A1", "Sheet1");
      expect(cell.value).toBe("Budget!$A$1");
      cell = await engine.getCell("A2", "Sheet1");
      expect(cell.value).toBe("[book.xlsx]Budget");

      let payload: any = null;
      const unsub = doc.on("change", (p: any) => {
        payload = p;
      });
      doc.renameSheet("sheet_2", "Data");
      unsub();

      expect(payload?.recalc).toBe(false);
      expect(Array.isArray(payload?.sheetMetaDeltas)).toBe(true);
      expect(payload.sheetMetaDeltas.length).toBeGreaterThan(0);

      await engineApplyDocumentChange(syncTarget, payload);

      cell = await engine.getCell("A1", "Sheet1");
      expect(cell.value).toBe("Data!$A$1");
      cell = await engine.getCell("A2", "Sheet1");
      expect(cell.value).toBe("[book.xlsx]Data");
    } finally {
      engine.terminate();
    }
  });
});

const previousSelf = (globalThis as any).self;
const workerGlobal = new MockWorkerGlobal();
// `engine.worker.ts` expects a WebWorker-like `self`.
(globalThis as any).self = workerGlobal;

let workerModulePromise: Promise<unknown> | null = null;
function loadWorkerModule(): Promise<unknown> {
  if (!workerModulePromise) {
    workerModulePromise = import("../engine.worker.ts");
  }
  return workerModulePromise;
}
