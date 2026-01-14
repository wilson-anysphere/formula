import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";
import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";
import type { InitMessage, RpcRequest, WorkerInboundMessage, WorkerOutboundMessage } from "../protocol.ts";

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

async function loadFormulaWasm() {
  const entry = formulaWasmNodeEntryUrl();
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports are
  // exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  const mod = await import(/* @vite-ignore */ entry);
  return (mod as any).default ?? mod;
}

class WasmBackedWorker implements WorkerLike {
  private readonly wasm: any;
  private port: MockMessagePort | null = null;
  private workbook: any | null = null;

  constructor(wasm: any) {
    this.wasm = wasm;
  }

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") {
      return;
    }

    this.port = init.port as unknown as MockMessagePort;
    this.workbook = new this.wasm.WasmWorkbook();

    this.port.addEventListener("message", (event) => {
      const msg = event.data as WorkerInboundMessage;
      if (!msg || typeof msg !== "object" || (msg as any).type !== "request") {
        return;
      }

      const req = msg as RpcRequest;
      const params = req.params as any;

      try {
        let result: unknown;
        switch (req.method) {
          case "newWorkbook":
            this.workbook = new this.wasm.WasmWorkbook();
            result = null;
            break;
          case "loadFromXlsxBytes":
            this.workbook = this.wasm.WasmWorkbook.fromXlsxBytes(params.bytes);
            result = null;
            break;
          case "loadFromEncryptedXlsxBytes":
            this.workbook = this.wasm.WasmWorkbook.fromEncryptedXlsxBytes(params.bytes, params.password);
            result = null;
            break;
          case "recalculate":
            result = this.workbook?.recalculate(params.sheet);
            break;
          case "getCell":
            result = this.workbook?.getCell(params.address, params.sheet);
            break;
          default:
            throw new Error(`unsupported method: ${req.method}`);
        }

        const response: WorkerOutboundMessage = { type: "response", id: req.id, ok: true, result };
        this.port?.postMessage(response);
      } catch (err) {
        const response: WorkerOutboundMessage = {
          type: "response",
          id: req.id,
          ok: false,
          error: err instanceof Error ? err.message : String(err),
        };
        this.port?.postMessage(response);
      }
    });

    const ready: WorkerOutboundMessage = { type: "ready" };
    this.port.postMessage(ready);
  }

  terminate(): void {
    this.port?.close();
    this.port = null;
    this.workbook = null;
  }
}

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

describeWasm("EngineWorker encrypted workbook load (wasm)", () => {
  it("decrypts and opens an encrypted XLSX fixture via loadWorkbookFromEncryptedXlsxBytes", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const repoRoot = path.resolve(__dirname, "../../../..");
    const encryptedPath = path.join(repoRoot, "fixtures", "encrypted", "ooxml", "agile-empty-password.xlsx");
    const bytes = new Uint8Array(readFileSync(encryptedPath));

    try {
      await engine.loadWorkbookFromEncryptedXlsxBytes(bytes, "");
      await engine.recalculate();

      const a1 = (await engine.getCell("A1", "Sheet1")) as any;
      const b1 = (await engine.getCell("B1", "Sheet1")) as any;
      expect(a1.value).toBe(1);
      expect(b1.value).toBe("Hello");
    } finally {
      engine.terminate();
    }
  });

  it("decrypts and opens an encrypted XLSM fixture via loadWorkbookFromEncryptedXlsxBytes", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const repoRoot = path.resolve(__dirname, "../../../..");
    const encryptedPath = path.join(repoRoot, "fixtures", "encrypted", "ooxml", "agile-basic.xlsm");
    const bytes = new Uint8Array(readFileSync(encryptedPath));

    try {
      await engine.loadWorkbookFromEncryptedXlsxBytes(bytes, "password");
      await engine.recalculate();

      const a1 = (await engine.getCell("A1", "Sheet1")) as any;
      expect(a1.value).toBe(null);
    } finally {
      engine.terminate();
    }
  });

  it("decrypts and opens an encrypted XLSB fixture via loadWorkbookFromEncryptedXlsxBytes", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const repoRoot = path.resolve(__dirname, "../../../..");
    const encryptedPath = path.join(repoRoot, "fixtures", "encrypted", "encrypted.xlsb");
    const bytes = new Uint8Array(readFileSync(encryptedPath));

    try {
      await engine.loadWorkbookFromEncryptedXlsxBytes(bytes, "tika");
      await engine.recalculate();

      const a1 = (await engine.getCell("A1", "Sheet1")) as any;
      const b1 = (await engine.getCell("B1", "Sheet1")) as any;
      expect(a1.value).toBe("You can't see me");
      expect(b1.value).toBe(null);
    } finally {
      engine.terminate();
    }
  });

  it("rejects invalid passwords with a clear error message", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel,
    });

    const __filename = fileURLToPath(import.meta.url);
    const __dirname = path.dirname(__filename);
    const repoRoot = path.resolve(__dirname, "../../../..");
    const encryptedPath = path.join(repoRoot, "fixtures", "encrypted", "ooxml", "agile-empty-password.xlsx");
    const bytes = new Uint8Array(readFileSync(encryptedPath));

    try {
      await expect(engine.loadWorkbookFromEncryptedXlsxBytes(bytes, "wrong-password")).rejects.toThrow(/invalid password/i);
    } finally {
      engine.terminate();
    }
  });
});
