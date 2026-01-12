import type { MacroPermission, MacroPermissionRequest, MacroSecurityStatus, MacroTrustDecision } from "./types";
import * as nativeDialogs from "../tauri/nativeDialogs.js";
import { showQuickPick } from "../extensions/ui.js";

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
 * A minimal built-in controller that uses non-blocking dialogs.
 * Desktop builds should replace this with an app-native modal.
 */
export class DefaultMacroSecurityController implements MacroSecurityController {
  async requestTrustDecision(prompt: MacroTrustDecisionPrompt): Promise<MacroTrustDecision | null> {
    const workbook = describeWorkbook(prompt.status, prompt.workbookId);
    const signature = describeSignature(prompt.status);
    const decision = await showQuickPick<MacroTrustDecision>(
      [
        { label: "Blocked", value: "blocked", description: "Do not run macros" },
        { label: "Trust once", value: "trusted_once", description: "Allow macros for this run only" },
        { label: "Trust always", value: "trusted_always", description: "Always allow macros for this workbook" },
        { label: "Trust signed only", value: "trusted_signed_only", description: "Only allow signed macros" },
      ],
      {
        placeHolder: `Macro Trust Center: ${prompt.macroId} (${workbook}) · signature ${signature}`,
      },
    );
    return decision;
  }

  async requestPermissions(prompt: MacroPermissionPrompt): Promise<MacroPermission[] | null> {
    const req = prompt.request;
    const workbook = req.workbookOriginPath ?? prompt.workbookId;
    const requested = Array.from(new Set(req.requested ?? []));
    if (requested.length === 0) return [];

    const alreadyGranted = prompt.alreadyGranted.length ? `\nAlready granted: ${prompt.alreadyGranted.join(", ")}` : "";

    const ok = await nativeDialogs.confirm(
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
