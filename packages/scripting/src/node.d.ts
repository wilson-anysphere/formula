import type { Workbook } from "./index";

export interface ScriptPermissions {
  network?: "none" | "allowlist" | "full";
  networkAllowlist?: string[];
}

export interface ScriptConsoleEntry {
  level: "log" | "info" | "warn" | "error";
  message: string;
}

export interface ScriptErrorShape {
  name?: string;
  message: string;
  stack?: string;
}

export interface ScriptRunResult {
  logs: ScriptConsoleEntry[];
  error?: ScriptErrorShape;
}

export class ScriptRuntime {
  constructor(workbook: Workbook);
  run(code: string, options?: { permissions?: ScriptPermissions; timeoutMs?: number }): Promise<ScriptRunResult>;
}

export * from "./index";

