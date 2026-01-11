export type MacroLanguage = "vba" | "typescript" | "python";

export interface MacroInfo {
  /** Workbook-scoped unique id. */
  id: string;
  /** Display name, e.g. "Module1.Macro1" or "Macro1". */
  name: string;
  language: MacroLanguage;
  /** Optional source location for debugging. */
  module?: string;
}

/** Trust Center decision for a workbook's macro fingerprint. */
export type MacroTrustDecision =
  | "blocked"
  | "trusted_once"
  | "trusted_always"
  | "trusted_signed_only";

/**
 * Digital signature state for the workbook's VBA project.
 *
 * These values intentionally mirror the backend's snake_case strings so the
 * frontend can round-trip them back to the Trust Center APIs.
 *
 * Note: `signed_untrusted` is reserved for a future state where the signature is
 * cryptographically valid but the certificate chain is not trusted.
 */
export type MacroSignatureStatus =
  | "unsigned"
  | "signed_unverified"
  | "signed_parse_error"
  | "signed_verified"
  | "signed_invalid"
  | "signed_untrusted";

export interface MacroSignatureInfo {
  status: MacroSignatureStatus;
  signerSubject?: string;
  signatureBase64?: string;
}

export interface MacroSecurityStatus {
  hasMacros: boolean;
  originPath?: string;
  workbookFingerprint?: string;
  signature?: MacroSignatureInfo;
  trust: MacroTrustDecision;
}

export type MacroBlockedReason = "not_trusted" | "signature_required";

export interface MacroBlockedError {
  reason: MacroBlockedReason;
  status: MacroSecurityStatus;
}

export interface MacroPermissionRequest {
  reason: string;
  macroId: string;
  workbookOriginPath?: string;
  requested: MacroPermission[];
}

export type MacroPermission =
  | "filesystem_read"
  | "filesystem_write"
  | "network"
  | "object_creation";

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

export interface MacroCellUpdate {
  sheetId: string;
  row: number;
  col: number;
  value: unknown | null;
  formula: string | null;
  displayValue: string;
}

export interface MacroRunResult {
  ok: boolean;
  output: string[];
  /**
   * Optional cell updates produced by the backend (for UIs that apply changes
   * incrementally).
   */
  updates?: MacroCellUpdate[];
  /**
   * Optional permission escalation request (e.g. sandbox denied an operation).
   * Present when `ok=false`.
   */
  permissionRequest?: MacroPermissionRequest;
  error?: {
    message: string;
    stack?: string;
    code?: string;
    blocked?: MacroBlockedError;
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
  getMacroSecurityStatus(workbookId: string): Promise<MacroSecurityStatus>;
  setMacroTrust(workbookId: string, decision: MacroTrustDecision): Promise<MacroSecurityStatus>;
  runMacro(request: MacroRunRequest): Promise<MacroRunResult>;
}
