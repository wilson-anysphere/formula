export type CellScalar = number | string | boolean | null;

export interface CellData {
  sheet: string;
  address: string;
  input: CellScalar;
  value: CellScalar;
}

export interface CellChange {
  sheet: string;
  address: string;
  value: CellScalar;
}

export interface RpcOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}

export interface InitMessage {
  type: "init";
  port: MessagePort;
  /**
   * URL to the WASM module (typically the wasm-bindgen JS glue entrypoint).
   *
   * Pass an empty string to disable WASM loading; RPC requests will fail with
   * `worker not initialized`.
   */
  wasmModuleUrl: string;
}

export interface ReadyMessage {
  type: "ready";
}

export interface RpcRequest {
  type: "request";
  id: number;
  method: string;
  params: unknown;
}

export interface RpcCancel {
  type: "cancel";
  id: number;
}

export interface RpcResponseOk {
  type: "response";
  id: number;
  ok: true;
  result: unknown;
}

export interface RpcResponseErr {
  type: "response";
  id: number;
  ok: false;
  error: string;
}

export type WorkerInboundMessage = RpcRequest | RpcCancel;
export type WorkerOutboundMessage = ReadyMessage | RpcResponseOk | RpcResponseErr;
