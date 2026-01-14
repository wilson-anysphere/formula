import { afterAll, describe, expect, it } from "vitest";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";

import { engineApplyDocumentChange, engineHydrateFromDocument, type EngineSyncTarget } from "../documentControllerSync.ts";
import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

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
  private listeners = new Set<(event: MessageEvent<unknown>) => void>();
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

  addEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.add(listener);
  }

  removeEventListener(_type: "message", listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.delete(listener);
  }

  private dispatchMessage(data: unknown): void {
    const event = { data } as MessageEvent<unknown>;
    this.onmessage?.(event);
    for (const listener of this.listeners) {
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

describeWasm("DocumentController range-run formatting → worker RPC → CELL() (wasm)", () => {
  afterAll(() => {
    (globalThis as any).self = previousSelf;
  });

  it('propagates large-range formatting (rangeRunDeltas) so CELL("protect") reflects range-run formatting', async () => {
    const wasmModuleUrl = new URL("./fixtures/formulaWasmNodeWrapper.mjs", import.meta.url).href;
    await loadWorkerModule();

    const engine = await EngineWorker.connect({
      worker: new LocalWorker(),
      wasmModuleUrl,
      channelFactory: createMockChannel,
    });

    // `engineHydrateFromDocument` / `engineApplyDocumentChange` operate on the narrower
    // `EngineSyncTarget` surface (sheet-first signatures for some metadata APIs). The public
    // EngineWorker client is an EngineClient-style API (sheet-last helpers like
    // `setCellStyleId(address, styleId, sheet)` and `setColWidth(col, width, sheet)`).
    //
    // Wrap it so `tsc` typechecking stays strict while still exercising the real worker RPC.
    const syncTarget: EngineSyncTarget = {
      loadWorkbookFromJson: (json) => engine.loadWorkbookFromJson(json),
      setCell: (address, value, sheet) => engine.setCell(address, value, sheet),
      setCells: (updates) => engine.setCells(updates),
      recalculate: (sheet) => engine.recalculate(sheet),
      setSheetDisplayName: (sheetId, name) => engine.setSheetDisplayName(sheetId, name),
      internStyle: (styleObj) => engine.internStyle(styleObj as any),
      setCellStyleId: (sheet, address, styleId) => engine.setCellStyleId(address, styleId, sheet),
      setFormatRunsByCol: (sheet, col, runs) => engine.setFormatRunsByCol(sheet, col, runs),
      setColWidthChars: (sheet, col, widthChars) => engine.setColWidthChars(sheet, col, widthChars),
    };

    try {
      const doc = new DocumentController();

      doc.setCellValue("Sheet1", "A1", "x");
      // Keep the formula cell out of the formatted rectangle so only the referenced cell's format changes.
      doc.setCellFormula("Sheet1", "AA1", 'CELL("protect",A1)');

      await engineHydrateFromDocument(syncTarget, doc);

      let cell = await engine.getCell("AA1", "Sheet1");
      expect(cell.value).toBe(1);

      // Apply a large-formatting patch so DocumentController uses compressed range-run formatting.
      let payload: any = null;
      const unsub = doc.on("change", (p: any) => {
        payload = p;
      });
      doc.setRangeFormat("Sheet1", "A1:Z2000", { protection: { locked: false } });
      unsub();

      expect(payload?.recalc).toBe(false);
      expect(Array.isArray(payload?.rangeRunDeltas)).toBe(true);
      expect(payload.rangeRunDeltas.length).toBeGreaterThan(0);

      await engineApplyDocumentChange(syncTarget, payload, { getStyleById: (id) => doc.styleTable.get(id) });

      cell = await engine.getCell("AA1", "Sheet1");
      expect(cell.value).toBe(0);

      // Clear formatting back to default by passing `null` (DocumentController semantics: clear style).
      payload = null;
      const unsub2 = doc.on("change", (p: any) => {
        payload = p;
      });
      doc.setRangeFormat("Sheet1", "A1:Z2000", null);
      unsub2();

      expect(payload?.recalc).toBe(false);
      expect(Array.isArray(payload?.rangeRunDeltas)).toBe(true);
      expect(payload.rangeRunDeltas.length).toBeGreaterThan(0);

      await engineApplyDocumentChange(syncTarget, payload, { getStyleById: (id) => doc.styleTable.get(id) });

      cell = await engine.getCell("AA1", "Sheet1");
      expect(cell.value).toBe(1);
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
