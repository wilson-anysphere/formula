import type { MacroPermission, MacroPermissionRequest, MacroSecurityStatus, MacroTrustDecision } from "./types";

export interface MacroTrustDecisionPrompt {
  workbookId: string;
  macroId: string;
  status: MacroSecurityStatus;
}

export interface MacroPermissionPrompt {
  workbookId: string;
  request: MacroPermissionRequest;
  alreadyGranted: MacroPermission[];
}

/**
 * UI-facing macro security controller.
 *
 * Desktop builds should implement this with an app-native modal that matches the
 * backend's Trust Center + sandbox permission model.
 */
export interface MacroSecurityController {
  /**
   * Prompt the user for a Trust Center decision (blocked/trusted once/always/signed only).
   *
   * Returning `null` indicates the user cancelled the prompt.
   */
  requestTrustDecision(prompt: MacroTrustDecisionPrompt): Promise<MacroTrustDecision | null>;

  /**
   * Prompt the user to grant additional permissions for the current macro run.
   *
   * Returning `null` indicates the user declined or cancelled.
   */
  requestPermissions(prompt: MacroPermissionPrompt): Promise<MacroPermission[] | null>;
}

function describeWorkbook(status: MacroSecurityStatus, workbookId: string): string {
  return status.originPath ?? workbookId;
}

function describeSignature(status: MacroSecurityStatus): string {
  const sig = status.signature;
  if (!sig) return "unknown";
  const suffix = sig.signerSubject ? ` (${sig.signerSubject})` : "";
  return `${sig.status}${suffix}`;
}

/**
 * A minimal built-in controller that uses `window.confirm` / `window.prompt`.
 * Desktop builds should replace this with an app-native modal.
 */
export class DefaultMacroSecurityController implements MacroSecurityController {
  async requestTrustDecision(prompt: MacroTrustDecisionPrompt): Promise<MacroTrustDecision | null> {
    const workbook = describeWorkbook(prompt.status, prompt.workbookId);
    const signature = describeSignature(prompt.status);
    const message =
      `This workbook contains VBA macros.\n\n` +
      `Workbook: ${workbook}\n` +
      `Macro: ${prompt.macroId}\n` +
      `Signature: ${signature}\n` +
      `Current Trust Center decision: ${prompt.status.trust}\n\n` +
      `Choose a Trust Center decision:\n` +
      `  1) blocked\n` +
      `  2) trusted_once\n` +
      `  3) trusted_always\n` +
      `  4) trusted_signed_only\n\n` +
      `Enter 1-4:`;

    const input = window.prompt(message, "2");
    if (input == null) return null;
    const value = input.trim();
    switch (value) {
      case "1":
      case "blocked":
        return "blocked";
      case "2":
      case "trusted_once":
        return "trusted_once";
      case "3":
      case "trusted_always":
        return "trusted_always";
      case "4":
      case "trusted_signed_only":
        return "trusted_signed_only";
      default:
        return null;
    }
  }

  async requestPermissions(prompt: MacroPermissionPrompt): Promise<MacroPermission[] | null> {
    const req = prompt.request;
    const workbook = req.workbookOriginPath ?? prompt.workbookId;
    const requested = Array.from(new Set(req.requested ?? []));
    if (requested.length === 0) return [];

    const alreadyGranted = prompt.alreadyGranted.length ? `\nAlready granted: ${prompt.alreadyGranted.join(", ")}` : "";

    const ok = window.confirm(
      `Macro permission request\n\n` +
        `Workbook: ${workbook}\n` +
        `Macro: ${req.macroId}\n` +
        `Requested: ${requested.join(", ")}${alreadyGranted}\n\n` +
        `Reason: ${req.reason}\n\n` +
        `Grant these permissions for this run?`
    );
    return ok ? requested : null;
  }
}
