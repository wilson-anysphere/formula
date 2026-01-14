import { describe, expect, it } from "vitest";

import { formulaWasmNodeEntryUrl } from "../../../../scripts/build-formula-wasm-node.mjs";

import { EngineWorker, type MessageChannelLike, type WorkerLike } from "../worker/EngineWorker.ts";
import type {
  InitMessage,
  RpcRequest,
  WorkerInboundMessage,
  WorkerOutboundMessage
} from "../protocol.ts";

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
  // wasm-pack `--target nodejs` outputs CommonJS. Under ESM dynamic import, the exports
  // are exposed on `default`.
  // eslint-disable-next-line @typescript-eslint/ban-ts-comment
  // @ts-ignore - `@vite-ignore` is required for runtime-defined file URLs.
  const mod = await import(/* @vite-ignore */ entry);
  return (mod as any).default ?? mod;
}

class WasmBackedWorker implements WorkerLike {
  private readonly wasm: any;
  private port: MockMessagePort | null = null;

  constructor(wasm: any) {
    this.wasm = wasm;
  }

  postMessage(message: unknown): void {
    const init = message as InitMessage;
    if (!init || typeof init !== "object" || (init as any).type !== "init") {
      return;
    }

    this.port = init.port as unknown as MockMessagePort;

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
          case "supportedLocaleIds":
            result = this.wasm.supportedLocaleIds();
            break;
          case "getLocaleInfo":
            result = this.wasm.getLocaleInfo(params.localeId);
            break;
          case "lexFormula":
            result = this.wasm.lexFormula(params.formula, params.options);
            break;
          case "lexFormulaPartial":
            result = this.wasm.lexFormulaPartial(params.formula, params.options);
            break;
          case "parseFormulaPartial":
            result = this.wasm.parseFormulaPartial(params.formula, params.cursor, params.options);
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
          error: err instanceof Error ? err.message : String(err)
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
  }
}

const skipWasmBuild = process.env.FORMULA_SKIP_WASM_BUILD === "1" || process.env.FORMULA_SKIP_WASM_BUILD === "true";
const describeWasm = skipWasmBuild ? describe.skip : describe;

describeWasm("EngineWorker editor tooling RPCs (wasm)", () => {
  it("supportedLocaleIds returns the engine-supported locale id list", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const ids = await engine.supportedLocaleIds();
      expect(ids).toEqual([...ids].sort());
      expect(ids).toContain("en-US");
      expect(ids).toContain("de-DE");
      expect(ids).toContain("fr-FR");
      expect(ids).toContain("es-ES");
    } finally {
      engine.terminate();
    }
  });

  it("getLocaleInfo returns formula punctuation and boolean literals", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const info = await engine.getLocaleInfo("de-DE");
      expect(info).toEqual({
        localeId: "de-DE",
        decimalSeparator: ",",
        argSeparator: ";",
        arrayRowSeparator: ";",
        arrayColSeparator: "\\",
        thousandsSeparator: ".",
        isRtl: false,
        booleanTrue: "WAHR",
        booleanFalse: "FALSCH"
      });
    } finally {
      engine.terminate();
    }
  });

  it("lexFormula returns token DTOs with UTF-16 spans", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      expect(await engine.lexFormula("=1+2")).toEqual([
        { kind: "Number", span: { start: 1, end: 2 }, value: "1" },
        { kind: "Plus", span: { start: 2, end: 3 } },
        { kind: "Number", span: { start: 3, end: 4 }, value: "2" },
        { kind: "Eof", span: { start: 4, end: 4 } }
      ]);
    } finally {
      engine.terminate();
    }
  });

  it("lexFormula spans use UTF-16 code unit indexing (surrogate pairs)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const formula = '=\"😀\"+1';
      expect(formula.length).toBe(7); // `😀` is 2 UTF-16 code units.
      const tokens = await engine.lexFormula(formula);
      expect(tokens[0]).toEqual({ kind: "String", span: { start: 1, end: 5 }, value: "😀" });
    } finally {
      engine.terminate();
    }
  });

  it("parseFormulaPartial returns function call context", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const result = await engine.parseFormulaPartial("=SUM(1,", 6);
      expect(result.context.function).toEqual({ name: "SUM", argIndex: 0 });
      expect(result.error?.span).toEqual({ start: 6, end: 6 });
    } finally {
      engine.terminate();
    }
  });

  it("parseFormulaPartial strips _xlfn. prefix from function context names", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const formula = "=_xlfn.SEQUENCE(1,";
      const cursor = formula.length;
      const result = await engine.parseFormulaPartial(formula, cursor);
      expect(result.context.function).toEqual({ name: "SEQUENCE", argIndex: 1 });
      expect(result.error?.span).toEqual({ start: cursor, end: cursor });

      const localized = "=_xlfn.SEQUENZ(1;";
      const localizedCursor = localized.length;
      const localizedResult = await engine.parseFormulaPartial(localized, localizedCursor, { localeId: "de-DE" });
      expect(localizedResult.context.function).toEqual({ name: "SEQUENCE", argIndex: 1 });
      expect(localizedResult.error?.span).toEqual({ start: localizedCursor, end: localizedCursor });
    } finally {
      engine.terminate();
    }
  });

  it("parseFormulaPartial cursor is UTF-16 indexed (surrogate pairs before cursor)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const formula = '=\"😀\"&SUM(1,';
      // Ensure we pass a UTF-16 code unit cursor (JS string indexing).
      const cursor = formula.indexOf(",") + 1;
      const result = await engine.parseFormulaPartial(formula, cursor);
      expect(result.context.function).toEqual({ name: "SUM", argIndex: 1 });
      expect(result.error?.span).toEqual({ start: cursor, end: cursor });
    } finally {
      engine.terminate();
    }
  });

  it("lexFormula honors localeId options (argument separator)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const tokens = await engine.lexFormula("=SUMME(1;2)", { localeId: "de-DE" });
      expect(tokens.some((t) => t.kind === "ArgSep")).toBe(true);

      await expect(engine.lexFormula("=SUMME(1;2)")).rejects.toThrow(/Unexpected character `;`/);
    } finally {
      engine.terminate();
    }
  });

  it("lexFormula honors referenceStyle options (R1C1 vs A1)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const tokens = await engine.lexFormula("=R1C1", { referenceStyle: "R1C1" });
      expect(tokens.some((t) => t.kind === "R1C1Cell")).toBe(true);

      const defaultTokens = await engine.lexFormula("=R1C1");
      expect(defaultTokens.some((t) => t.kind === "R1C1Cell")).toBe(false);
    } finally {
      engine.terminate();
    }
  });

  it("parseFormulaPartial honors localeId options (argument separator)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const localized = await engine.parseFormulaPartial("=SUMME(1;2)", undefined, { localeId: "de-DE" });
      expect(localized.error).toBeNull();

      const defaultResult = await engine.parseFormulaPartial("=SUMME(1;2)");
      expect(defaultResult.error?.message).toContain("Unexpected character `;`");
      expect(defaultResult.error?.span).toEqual({ start: 8, end: 9 });
    } finally {
      engine.terminate();
    }
  });

  it("parseFormulaPartial honors referenceStyle options (R1C1 vs A1)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const r1c1 = await engine.parseFormulaPartial("=R1C1", { referenceStyle: "R1C1" });
      expect(r1c1.error).toBeNull();
      expect((r1c1.ast as any)?.expr?.CellRef).toBeTruthy();

      const defaultResult = await engine.parseFormulaPartial("=R1C1");
      expect(defaultResult.error).toBeNull();
      expect((defaultResult.ast as any)?.expr?.NameRef).toBeTruthy();
    } finally {
      engine.terminate();
    }
  });

  it("accepts full ParseOptions objects (snake_case) for backward compatibility", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const fullOpts = {
        locale: {
          decimal_separator: ".",
          arg_separator: ",",
          array_col_separator: ",",
          array_row_separator: ";",
          thousands_separator: null
        },
        reference_style: "R1C1",
        normalize_relative_to: null
      };

      const tokens = await engine.lexFormula("=R1C1", fullOpts as any);
      expect(tokens.some((t) => t.kind === "R1C1Cell")).toBe(true);

      const parsed = await engine.parseFormulaPartial("=R1C1", fullOpts as any);
      expect(parsed.error).toBeNull();
      expect((parsed.ast as any)?.expr?.CellRef).toBeTruthy();
    } finally {
      engine.terminate();
    }
  });

  it("surfaces a clear error when the options object has an unexpected shape", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      // Common mistake: wrong casing on `localeId`.
      await expect(engine.lexFormula("=1+2", { localeID: "de-DE" } as any)).rejects.toThrow(
        /options must be \{ localeId\?: string, referenceStyle\?:/
      );
      // Legacy call sites use `(formula, undefined, options)` when they want to supply a cursor
      // later. Ensure malformed objects aren't silently dropped by the overload parsing.
      await expect(engine.parseFormulaPartial("=1+2", undefined, { localeID: "de-DE" } as any)).rejects.toThrow(
        /options must be \{ localeId\?: string, referenceStyle\?:/
      );
    } finally {
      engine.terminate();
    }
  });

  it("surfaces unknown localeId errors from the WASM parser helpers", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      await expect(engine.lexFormula("=1+2", { localeId: "xx-XX" })).rejects.toThrow(
        /unknown localeId: xx-XX.*Supported locale ids/
      );
      await expect(engine.parseFormulaPartial("=1+2", undefined, { localeId: "xx-XX" })).rejects.toThrow(
        /unknown localeId: xx-XX.*Supported locale ids/
      );
    } finally {
      engine.terminate();
    }
  });

  it("lexFormulaPartial returns tokens + error for unterminated strings (best-effort)", async () => {
    const wasm = await loadFormulaWasm();
    const worker = new WasmBackedWorker(wasm);
    const engine = await EngineWorker.connect({
      worker,
      wasmModuleUrl: "mock://wasm",
      channelFactory: createMockChannel
    });

    try {
      const result = await engine.lexFormulaPartial("=\"hello");
      expect(result.tokens.length).toBeGreaterThan(0);
      expect(result.error?.message).toContain("Unterminated string literal");

      const stringToken = result.tokens.find((t) => t.kind === "String");
      expect(stringToken?.span).toEqual({ start: 1, end: 7 });
    } finally {
      engine.terminate();
    }
  });
});
