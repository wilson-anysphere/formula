import type {
  CellChange,
  CellData,
  CellScalar,
  EditOp,
  EditResult,
  FormulaPartialParseResult,
  FormulaParseOptions,
  FormulaToken,
  InitMessage,
  RpcCancel,
  RpcOptions,
  RpcRequest,
  RpcResponseErr,
  RpcResponseOk,
  WorkerOutboundMessage
} from "../protocol.ts";
import { isFormulaInput, normalizeFormulaText } from "../backend/formula.ts";

export interface WorkerLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  terminate(): void;
}

export interface MessagePortLike {
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start?(): void;
  close?(): void;
  addEventListener(type: "message", listener: (event: MessageEvent<unknown>) => void): void;
  removeEventListener(type: "message", listener: (event: MessageEvent<unknown>) => void): void;
}

export interface MessageChannelLike {
  port1: MessagePortLike;
  port2: MessagePort;
}

type PendingRequest = {
  resolve: (value: unknown) => void;
  reject: (err: unknown) => void;
  timeoutId?: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abortListener?: () => void;
};

type CellUpdate = { address: string; value: CellScalar; sheet?: string };

function normalizeCellScalar(value: CellScalar): CellScalar {
  if (typeof value !== "string") return value;
  if (!isFormulaInput(value)) return value;
  return normalizeFormulaText(value);
}

/**
 * Worker-backed Engine client using a MessagePort RPC transport.
 *
 * This is intentionally similar to the proven implementation in
 * `apps/desktop/src/engine/worker/EngineWorker.ts` so desktop and web share the
 * same transport behavior.
 */
export class EngineWorker {
  private readonly worker: WorkerLike;
  private readonly port: MessagePortLike;
  private readonly pending = new Map<number, PendingRequest>();
  private nextId = 1;

  private pendingCellUpdates: CellUpdate[] = [];
  private flushPromise: Promise<void> | null = null;

  private constructor(worker: WorkerLike, port: MessagePortLike) {
    this.worker = worker;
    this.port = port;

    const handler = (event: MessageEvent<unknown>) => {
      this.onMessage(event.data as WorkerOutboundMessage);
    };
    this.port.addEventListener("message", handler);
    this.port.start?.();
  }

  static async connect(options: {
    worker: WorkerLike;
    wasmModuleUrl?: string;
    wasmBinaryUrl?: string;
    channelFactory?: () => MessageChannelLike;
  }): Promise<EngineWorker> {
    const channel = options.channelFactory?.() ?? new MessageChannel();
    const port = channel.port1;
    const worker = options.worker;

    const ready = new Promise<void>((resolve) => {
      const onReady = (event: MessageEvent<unknown>) => {
        const msg = event.data as WorkerOutboundMessage;
        if (msg && typeof msg === "object" && (msg as any).type === "ready") {
          port.removeEventListener("message", onReady);
          resolve();
        }
      };
      port.addEventListener("message", onReady);
    });

    const initMessage: InitMessage = {
      type: "init",
      port: channel.port2,
      wasmModuleUrl: options.wasmModuleUrl ?? "",
      wasmBinaryUrl: options.wasmBinaryUrl
    };
    worker.postMessage(initMessage, [channel.port2]);

    const engine = new EngineWorker(worker, port);
    await ready;
    return engine;
  }

  terminate(): void {
    for (const [id, pending] of this.pending) {
      pending.reject(new Error(`worker terminated (request ${id})`));
    }
    this.pending.clear();
    this.worker.terminate();
    this.port.close?.();
  }

  async newWorkbook(options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("newWorkbook", {}, options);
  }

  async loadWorkbookFromJson(json: string, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("loadFromJson", { json }, options);
  }

  /**
   * Load a workbook from raw `.xlsx` bytes.
   *
   * Note: the payload is transferred to the worker to avoid an extra
   * structured-clone copy.
   *
   * - If `bytes` spans its entire backing buffer (`byteOffset === 0` and
   *   `byteLength === bytes.buffer.byteLength`), the buffer is transferred and
   *   detached on the calling side.
   * - If `bytes` is a view into a larger buffer, we first copy it to a compact
   *   `Uint8Array` so only the relevant range is transferred (leaving the
   *   original buffer intact).
   */
  async loadWorkbookFromXlsxBytes(bytes: Uint8Array, options?: RpcOptions): Promise<void> {
    await this.flush();
    let payload = bytes;
    if (payload.byteOffset !== 0 || payload.byteLength !== payload.buffer.byteLength) {
      payload = payload.slice();
    }
    await this.invoke("loadFromXlsxBytes", { bytes: payload }, options, [payload.buffer]);
  }

  async toJson(options?: RpcOptions): Promise<string> {
    await this.flush();
    return (await this.invoke("toJson", {}, options)) as string;
  }

  async getCell(
    address: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellData> {
    await this.flush();
    return (await this.invoke("getCell", { address, sheet }, options)) as CellData;
  }

  async getRange(
    range: string,
    sheet?: string,
    options?: RpcOptions
  ): Promise<CellData[][]> {
    await this.flush();
    return (await this.invoke("getRange", { range, sheet }, options)) as CellData[][];
  }

  async setCell(address: string, value: CellScalar, sheet?: string): Promise<void> {
    this.pendingCellUpdates.push({ address, value: normalizeCellScalar(value), sheet });
    await this.scheduleFlush();
  }

  async setCells(
    updates: Array<{ address: string; value: CellScalar; sheet?: string }>,
    options?: RpcOptions
  ): Promise<void> {
    if (updates.length === 0) {
      return;
    }
    await this.flush();
    const normalized = updates.map((update) => ({ ...update, value: normalizeCellScalar(update.value) }));
    await this.invoke("setCells", { updates: normalized }, options);
  }

  async setRange(
    range: string,
    values: CellScalar[][],
    sheet?: string,
    options?: RpcOptions
  ): Promise<void> {
    await this.flush();
    const normalizedValues = values.map((row) => row.map((value) => normalizeCellScalar(value)));
    await this.invoke("setRange", { range, values: normalizedValues, sheet }, options);
  }

  async setLocale(localeId: string, options?: RpcOptions): Promise<boolean> {
    await this.flush();
    return (await this.invoke("setLocale", { localeId }, options)) as boolean;
  }

  async recalculate(sheet?: string, options?: RpcOptions): Promise<CellChange[]> {
    await this.flush();
    return (await this.invoke("recalculate", { sheet }, options)) as CellChange[];
  }

  async applyOperation(op: EditOp, options?: RpcOptions): Promise<EditResult> {
    await this.flush();
    return (await this.invoke("applyOperation", { op }, options)) as EditResult;
  }

  /**
   * Tokenize a formula string for editor tooling (syntax highlighting, etc).
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async lexFormula(formula: string, options?: FormulaParseOptions, rpcOptions?: RpcOptions): Promise<FormulaToken[]> {
    return (await this.invoke("lexFormula", { formula, options }, rpcOptions)) as FormulaToken[];
  }

  /**
   * Best-effort partial parse for editor/autocomplete scenarios.
   *
   * Note: this RPC is independent of workbook state and intentionally does NOT
   * flush pending `setCell` batches.
   */
  async parseFormulaPartial(
    formula: string,
    cursor?: number,
    options?: FormulaParseOptions,
    rpcOptions?: RpcOptions
  ): Promise<FormulaPartialParseResult> {
    return (await this.invoke("parseFormulaPartial", { formula, cursor, options }, rpcOptions)) as FormulaPartialParseResult;
  }

  async setSheetDimensions(sheet: string, rows: number, cols: number, options?: RpcOptions): Promise<void> {
    await this.flush();
    await this.invoke("setSheetDimensions", { sheet, rows, cols }, options);
  }

  async getSheetDimensions(sheet: string, options?: RpcOptions): Promise<{ rows: number; cols: number }> {
    await this.flush();
    return (await this.invoke("getSheetDimensions", { sheet }, options)) as { rows: number; cols: number };
  }

  private async scheduleFlush(): Promise<void> {
    if (this.flushPromise) {
      return this.flushPromise;
    }
    if (this.pendingCellUpdates.length === 0) {
      return;
    }

    this.flushPromise = new Promise((resolve, reject) => {
      queueMicrotask(async () => {
        try {
          await this.flush();
          resolve();
        } catch (err) {
          reject(err);
        } finally {
          this.flushPromise = null;
        }
      });
    });

    return this.flushPromise;
  }

  private async flush(): Promise<void> {
    if (this.pendingCellUpdates.length === 0) {
      return;
    }
    const updates = this.pendingCellUpdates;
    this.pendingCellUpdates = [];
    await this.invoke("setCells", { updates });
  }

  private invoke(
    method: string,
    params: unknown,
    options?: RpcOptions,
    transfer?: Transferable[]
  ): Promise<unknown> {
    const id = this.nextId++;
    const request: RpcRequest = { type: "request", id, method, params };

    return new Promise((resolve, reject) => {
      const pending: PendingRequest = { resolve, reject };
      this.pending.set(id, pending);

      if (options?.timeoutMs != null) {
        pending.timeoutId = setTimeout(() => {
          this.cancel(id, new Error(`request timed out after ${options.timeoutMs}ms`));
        }, options.timeoutMs);
      }

      if (options?.signal) {
        pending.signal = options.signal;
        if (options.signal.aborted) {
          this.cancel(id, new Error("request aborted"));
          return;
        }
        pending.abortListener = () => this.cancel(id, new Error("request aborted"));
        options.signal.addEventListener("abort", pending.abortListener, { once: true });
      }

      try {
        if (transfer && transfer.length > 0) {
          this.port.postMessage(request, transfer);
        } else {
          this.port.postMessage(request);
        }
      } catch (err) {
        this.pending.delete(id);
        pending.timeoutId && clearTimeout(pending.timeoutId);
        if (pending.signal && pending.abortListener) {
          pending.signal.removeEventListener("abort", pending.abortListener);
        }
        reject(err);
      }
    });
  }

  private cancel(id: number, error: Error): void {
    const pending = this.pending.get(id);
    if (!pending) {
      return;
    }

    this.pending.delete(id);
    pending.timeoutId && clearTimeout(pending.timeoutId);
    if (pending.signal && pending.abortListener) {
      pending.signal.removeEventListener("abort", pending.abortListener);
    }

    const cancelMessage: RpcCancel = { type: "cancel", id };
    try {
      this.port.postMessage(cancelMessage);
    } catch {
      // Ignore; cancellation is best-effort and may race with worker teardown.
    }
    pending.reject(error);
  }

  private onMessage(msg: WorkerOutboundMessage): void {
    if (msg.type !== "response") {
      return;
    }

    const pending = this.pending.get(msg.id);
    if (!pending) {
      return;
    }

    this.pending.delete(msg.id);
    pending.timeoutId && clearTimeout(pending.timeoutId);
    if (pending.signal && pending.abortListener) {
      pending.signal.removeEventListener("abort", pending.abortListener);
    }

    if ((msg as RpcResponseOk).ok) {
      pending.resolve((msg as RpcResponseOk).result);
      return;
    }

    pending.reject(new Error((msg as RpcResponseErr).error));
  }
}
