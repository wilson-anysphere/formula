import type { CommandRegistry } from "../extensions/commandRegistry.js";
import { PanelIds } from "../panels/panelRegistry.js";

export type MacrosPanelFocusTarget =
  | "runner-select"
  | "runner-run"
  | "runner-trust-center"
  | "recorder-start"
  | "recorder-stop";

export const RIBBON_MACRO_COMMAND_IDS = [
  // View → Macros.
  "view.macros.viewMacros",
  "view.macros.viewMacros.run",
  "view.macros.viewMacros.edit",
  "view.macros.viewMacros.delete",
  "view.macros.recordMacro",
  "view.macros.recordMacro.stop",
  "view.macros.useRelativeReferences",

  // Developer → Code.
  "developer.code.visualBasic",
  "developer.code.macros",
  "developer.code.macros.run",
  "developer.code.macros.edit",
  "developer.code.recordMacro",
  "developer.code.recordMacro.stop",
  "developer.code.useRelativeReferences",
  "developer.code.macroSecurity",
  "developer.code.macroSecurity.trustCenter",
] as const;

export type RibbonMacroCommandId = (typeof RIBBON_MACRO_COMMAND_IDS)[number];

export type RibbonMacroCommandHandlers = {
  openPanel: (panelId: string) => void;
  focusScriptEditorPanel: () => void;
  focusVbaMigratePanel: () => void;
  setPendingMacrosPanelFocus: (target: MacrosPanelFocusTarget | null) => void;
  startMacroRecorder: () => void;
  stopMacroRecorder: () => void;
  isTauri: () => boolean;
};

export function registerRibbonMacroCommands(params: {
  commandRegistry: CommandRegistry;
  handlers: RibbonMacroCommandHandlers;
}): void {
  const { commandRegistry, handlers } = params;
  const {
    openPanel,
    focusScriptEditorPanel,
    focusVbaMigratePanel,
    setPendingMacrosPanelFocus,
    startMacroRecorder,
    stopMacroRecorder,
    isTauri,
  } = handlers;

  const openMacrosPanel = (): void => openPanel(PanelIds.MACROS);
  const openScriptEditor = (): void => {
    openPanel(PanelIds.SCRIPT_EDITOR);
    focusScriptEditorPanel();
  };

  const titleForCommand = (commandId: RibbonMacroCommandId): string => {
    switch (commandId) {
      case "view.macros.viewMacros":
        return "View Macros…";
      case "view.macros.viewMacros.run":
      case "developer.code.macros.run":
        return "Run Macro…";
      case "view.macros.viewMacros.edit":
      case "developer.code.macros.edit":
        return "Edit Macro…";
      case "view.macros.viewMacros.delete":
        return "Delete Macro…";
      case "view.macros.recordMacro":
      case "developer.code.recordMacro":
        return "Record Macro";
      case "view.macros.recordMacro.stop":
      case "developer.code.recordMacro.stop":
        return "Stop Recording";
      case "view.macros.useRelativeReferences":
      case "developer.code.useRelativeReferences":
        return "Use Relative References";
      case "developer.code.visualBasic":
        return "Visual Basic";
      case "developer.code.macros":
        return "Macros…";
      case "developer.code.macroSecurity":
        return "Macro Security";
      case "developer.code.macroSecurity.trustCenter":
        return "Trust Center";
      default:
        return commandId;
    }
  };

  const runCommand = (commandId: RibbonMacroCommandId): void => {
    switch (commandId) {
      case "view.macros.viewMacros":
      case "view.macros.viewMacros.run":
      case "view.macros.viewMacros.edit":
      case "view.macros.viewMacros.delete": {
        // Clear any previously-requested focus so that edit/Visual Basic actions don't
        // get focus stolen by an earlier async macro runner render.
        if (commandId.endsWith(".edit")) setPendingMacrosPanelFocus(null);
        if (commandId === "view.macros.viewMacros") setPendingMacrosPanelFocus("runner-select");
        if (commandId.endsWith(".run")) setPendingMacrosPanelFocus("runner-run");
        if (commandId.endsWith(".delete")) setPendingMacrosPanelFocus("runner-select");
        openMacrosPanel();
        // "Edit…" in Excel normally opens an editor; best-effort surface our Script Editor panel too.
        if (commandId.endsWith(".edit")) openScriptEditor();
        return;
      }

      case "view.macros.recordMacro":
        setPendingMacrosPanelFocus("recorder-stop");
        startMacroRecorder();
        openMacrosPanel();
        return;
      case "view.macros.recordMacro.stop":
        setPendingMacrosPanelFocus("recorder-start");
        stopMacroRecorder();
        openMacrosPanel();
        return;

      case "view.macros.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. This command is intentionally a no-op.
        return;

      case "developer.code.macros":
      case "developer.code.macros.run":
      case "developer.code.macros.edit": {
        // Clear any previously-requested focus so that edit/Visual Basic actions don't
        // get focus stolen by an earlier async macro runner render.
        if (commandId.endsWith(".edit")) setPendingMacrosPanelFocus(null);
        if (commandId === "developer.code.macros") setPendingMacrosPanelFocus("runner-select");
        if (commandId.endsWith(".run")) setPendingMacrosPanelFocus("runner-run");
        openMacrosPanel();
        if (commandId.endsWith(".edit")) openScriptEditor();
        return;
      }

      case "developer.code.macroSecurity":
      case "developer.code.macroSecurity.trustCenter":
        setPendingMacrosPanelFocus("runner-trust-center");
        openMacrosPanel();
        return;

      case "developer.code.recordMacro":
        setPendingMacrosPanelFocus("recorder-stop");
        startMacroRecorder();
        openMacrosPanel();
        return;
      case "developer.code.recordMacro.stop":
        setPendingMacrosPanelFocus("recorder-start");
        stopMacroRecorder();
        openMacrosPanel();
        return;

      case "developer.code.useRelativeReferences":
        // Toggle state is handled by the ribbon UI; we don't currently implement a
        // "relative reference" mode in the macro recorder. This command is intentionally a no-op.
        return;

      case "developer.code.visualBasic":
        setPendingMacrosPanelFocus(null);
        // Desktop builds expose a VBA migration panel (used as a stand-in for the VBA editor).
        if (isTauri()) {
          openPanel(PanelIds.VBA_MIGRATE);
          focusVbaMigratePanel();
        } else {
          openScriptEditor();
        }
        return;
      default:
        return;
    }
  };

  for (const commandId of RIBBON_MACRO_COMMAND_IDS) {
    commandRegistry.registerBuiltinCommand(commandId, titleForCommand(commandId), () => runCommand(commandId), {
      category: "Macros",
      icon: null,
      description: null,
      keywords: ["macros", "vba", "script"],
    });
  }
}
