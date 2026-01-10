export type MacroLanguage = "vba";

export interface MacroInfo {
  /** Workbook-scoped unique id. */
  id: string;
  /** Display name, e.g. "Module1.Macro1" or "Macro1". */
  name: string;
  language: MacroLanguage;
  /** Optional source location for debugging. */
  module?: string;
}

export type MacroPermission =
  | "filesystem_read"
  | "filesystem_write"
  | "network";

export interface MacroRunRequest {
  workbookId: string;
  macroId: string;
  /**
   * Sandboxed permissions to grant for this run.
   *
   * Default is `[]` (no filesystem/network). The host may further restrict.
   */
  permissions?: MacroPermission[];
  /**
   * Maximum execution time in milliseconds. Used to surface intent to the
   * backend; the backend should still enforce its own limits.
   */
  timeoutMs?: number;
}

export interface MacroRunResult {
  ok: boolean;
  output: string[];
  error?: {
    message: string;
    stack?: string;
  };
}

/**
 * Backend bridge that the desktop shell provides (e.g. via Tauri commands).
 *
 * The actual implementation lives outside this directory; the UI can remain
 * decoupled and testable by using this interface.
 */
export interface MacroBackend {
  listMacros(workbookId: string): Promise<MacroInfo[]>;
  runMacro(request: MacroRunRequest): Promise<MacroRunResult>;
}

