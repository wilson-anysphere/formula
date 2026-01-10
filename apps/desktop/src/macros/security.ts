import type { MacroPermission } from "./types";

export type MacroSecuritySetting = "disabled" | "prompt" | "enabled";

export interface MacroSecurityDecision {
  enabled: boolean;
  /**
   * If `enabled` is true, which permissions to grant. Default should be none.
   */
  permissions: MacroPermission[];
}

/**
 * UI-facing macro security controller.
 *
 * This is where we integrate with:
 * - Workbook-level "Enable macros" UX
 * - Task 32 permissions model (filesystem/network/etc)
 */
export interface MacroSecurityController {
  getSetting(workbookId: string): Promise<MacroSecuritySetting>;
  setSetting(workbookId: string, setting: MacroSecuritySetting): Promise<void>;

  /**
   * Prompt the user to enable macros (and optionally request additional
   * permissions for this run).
   */
  requestEnableMacros(workbookId: string): Promise<MacroSecurityDecision>;
}

/**
 * A minimal built-in controller that uses `window.confirm` / `window.prompt`.
 * Desktop builds should replace this with an app-native modal.
 */
export class DefaultMacroSecurityController implements MacroSecurityController {
  private settings = new Map<string, MacroSecuritySetting>();

  async getSetting(workbookId: string): Promise<MacroSecuritySetting> {
    return this.settings.get(workbookId) ?? "prompt";
  }

  async setSetting(workbookId: string, setting: MacroSecuritySetting): Promise<void> {
    this.settings.set(workbookId, setting);
  }

  async requestEnableMacros(workbookId: string): Promise<MacroSecurityDecision> {
    const enabled = window.confirm(
      "This workbook contains macros. Enable macros?\n\n" +
        "Macros can modify your workbook. Network and filesystem access is disabled by default."
    );
    if (!enabled) {
      return { enabled: false, permissions: [] };
    }

    // Keep permissions minimal by default; allow the user to opt into network.
    const allowNetwork = window.confirm("Allow this macro to access the network?");
    const permissions: MacroPermission[] = allowNetwork ? ["network"] : [];
    return { enabled: true, permissions };
  }
}

