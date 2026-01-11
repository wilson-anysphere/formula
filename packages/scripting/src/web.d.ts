import type { WorkbookLike } from "./index";

export interface ScriptPrincipal {
  type: string;
  id: string;
}

export interface PermissionSnapshot {
  filesystem?: { read?: string[]; readwrite?: string[] };
  network?: { mode?: "none" | "allowlist" | "full"; allowlist?: string[] };
  clipboard?: boolean;
  notifications?: boolean;
  automation?: boolean;
}

export interface ScriptConsoleEntry {
  level: "log" | "info" | "warn" | "error";
  message: string;
}

export interface ScriptErrorShape {
  name?: string;
  message: string;
  stack?: string;
  code?: string;
  principal?: any;
  request?: any;
  reason?: string;
}

export interface AuditSink {
  log(event: any): void;
}

export interface ScriptRunResult {
  logs: ScriptConsoleEntry[];
  audit: any[];
  error?: ScriptErrorShape;
}

export interface ScriptRuntimeRunOptions {
  timeoutMs?: number;
  memoryMb?: number;
  signal?: AbortSignal;
  principal?: ScriptPrincipal;
  permissions?: PermissionSnapshot;
  permissionManager?: { getSnapshot(principal: ScriptPrincipal): PermissionSnapshot };
  auditSink?: AuditSink;
}

export class ScriptRuntime {
  constructor(workbook: WorkbookLike);
  run(code: string, options?: ScriptRuntimeRunOptions): Promise<ScriptRunResult>;
}

export * from "./index";
