import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState";
import { panelRegistry } from "../../panels/panelRegistry";
import { registerDesktopCommands } from "../../commands/registerDesktopCommands";
import { registerFormatPainterCommand } from "../../commands/formatPainterCommand";
import { registerRibbonMacroCommands } from "../../commands/registerRibbonMacroCommands";
import { registerDataQueriesCommands } from "../../commands/registerDataQueriesCommands";

import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbonCommandRegistryDisabling";
import { defaultRibbonSchema, type RibbonSchema } from "../ribbonSchema";

function collectRibbonCommandIds(schema: RibbonSchema): string[] {
  const ids = new Set<string>();
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        ids.add(button.id);
        for (const item of button.menuItems ?? []) {
          ids.add(item.id);
        }
      }
    }
  }
  return [...ids].sort();
}

function collectRibbonDropdownTriggerIds(schema: RibbonSchema): Set<string> {
  const ids = new Set<string>();
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        const kind = button.kind ?? "button";
        if (kind === "dropdown" && (button.menuItems?.length ?? 0) > 0) {
          // Dropdown triggers with menus do not invoke `onCommand`; only their menu items do.
          ids.add(button.id);
        }
      }
    }
  }
  return ids;
}

describe("Ribbon command wiring coverage (Home → Font dropdowns)", () => {
  it("uses canonical `format.*` ids for Font dropdown menu items", () => {
    const ids = collectRibbonCommandIds(defaultRibbonSchema);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(ids).toContain("home.font.fillColor");
    expect(ids).toContain("home.font.fontColor");
    expect(ids).toContain("home.font.borders");
    expect(ids).toContain("home.font.clearFormatting");

    // Font dropdown menu items were historically wired via `home.font.*` prefixes in `main.ts`.
    // These actions are now canonical `format.*` commands so ribbon/command-palette/keybindings
    // share a single command surface.
    const legacyMenuItemPrefixes = [
      "home.font.fillColor.",
      "home.font.fontColor.",
      "home.font.borders.",
      "home.font.clearFormatting.",
    ] as const;

    const legacyMenuItemIds = ids.filter((id) => legacyMenuItemPrefixes.some((prefix) => id.startsWith(prefix)));
    expect(
      legacyMenuItemIds,
      `Legacy Home→Font menu item ids should not exist in the ribbon schema:\n${legacyMenuItemIds
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);

    // Representative new ids (the complete set is covered by CommandRegistry + ribbon schema tests).
    expect(ids).toContain("format.fillColor.none");
    expect(ids).toContain("format.fontColor.black");
    expect(ids).toContain("format.borders.top");
    expect(ids).toContain("format.clearFormats");
  });

  it("does not use `home.font.*` prefix parsing for font dropdown menu items", () => {
    const sources = [
      { label: "main.ts", path: fileURLToPath(new URL("../../main.ts", import.meta.url)) },
      { label: "ribbon/commandHandlers.ts", path: fileURLToPath(new URL("../commandHandlers.ts", import.meta.url)) },
    ];

    for (const { label, path } of sources) {
      const source = readFileSync(path, "utf8");

      // Ensure the old prefix-parsing blocks were removed. (The dropdown trigger ids
      // like `home.font.fillColor` may still exist as fallbacks; only the menu item
      // prefix parsing is disallowed.)
      expect(source, `${label} should not include legacy Home→Font menu id prefixes`).not.toContain("home.font.fillColor.");
      expect(source, `${label} should not include legacy Home→Font menu id prefixes`).not.toContain("home.font.fontColor.");
      expect(source, `${label} should not include legacy Home→Font menu id prefixes`).not.toContain("home.font.borders.");
      expect(source, `${label} should not include legacy Home→Font menu id prefixes`).not.toContain("home.font.clearFormatting.");
    }
  });
});

const IMPLEMENTED_COMMAND_PREFIXES = [
  // Handled by `handleRibbonCommand` prefix logic (see `apps/desktop/src/main.ts`).
  "home.styles.cellStyles.",
  "home.styles.formatAsTable.",
];

function extractImplementedCommandIdsFromMainTs(schemaCommandIds: Set<string>): string[] {
  const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
  const source = readFileSync(mainTsPath, "utf8");
  const ids = new Set<string>();

  const addIfSchema = (id: string) => {
    if (schemaCommandIds.has(id)) ids.add(id);
  };

  for (const match of source.matchAll(/case\s+["']([^"']+)["']/g)) {
    addIfSchema(match[1]!);
  }

  for (const match of source.matchAll(/commandId\s*===\s*["']([^"']+)["']/g)) {
    addIfSchema(match[1]!);
  }

  // Keys in the `createRibbonActionsFromCommands({ ... })` overrides (commandOverrides/toggleOverrides).
  //
  // Use a structural regex anchored on the surrounding property names so we don't accidentally
  // treat other large object literals (e.g. `disabledById`) as command handlers.
  const extractOverrideKeys = (re: RegExp): void => {
    const match = re.exec(source);
    const body = match?.[1];
    if (!body) return;
    for (const keyMatch of body.matchAll(/["']([^"']+)["']\s*:/g)) {
      addIfSchema(keyMatch[1]!);
    }
  };
  extractOverrideKeys(/commandOverrides\s*:\s*\{([\s\S]*?)\}\s*,\s*onBeforeExecuteCommand\s*:/);
  extractOverrideKeys(/toggleOverrides\s*:\s*\{([\s\S]*?)\}\s*,\s*onUnknownCommand\s*:/);

  const presentPrefixes = IMPLEMENTED_COMMAND_PREFIXES.filter((prefix) => source.includes(prefix));
  for (const id of schemaCommandIds) {
    if (presentPrefixes.some((prefix) => id.startsWith(prefix))) {
      ids.add(id);
    }
  }

  return Array.from(ids).sort();
}

function registerCommandsForRibbonDisablingTest(commandRegistry: CommandRegistry): void {
  const layoutController = {
    layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
    openPanel(panelId: string) {
      this.layout = openPanel(this.layout, panelId, { panelRegistry });
    },
    closePanel(panelId: string) {
      this.layout = closePanel(this.layout, panelId);
    },
  } as any;

  registerDesktopCommands({
    commandRegistry,
    app: {} as any,
    layoutController,
    themeController: { setThemePreference: () => {} } as any,
    refreshRibbonUiState: () => {},
    applyFormattingToSelection: () => {},
    getActiveCellNumberFormat: () => null,
    getActiveCellIndentLevel: () => 0,
    openFormatCells: () => {},
    showQuickPick: async () => null,
    pageLayoutHandlers: {
      openPageSetupDialog: () => {},
      updatePageSetup: () => {},
      setPrintArea: () => {},
      clearPrintArea: () => {},
      addToPrintArea: () => {},
      exportPdf: () => {},
    },
    findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
    workbenchFileHandlers: {
      newWorkbook: () => {},
      openWorkbook: () => {},
      saveWorkbook: () => {},
      saveWorkbookAs: () => {},
      setAutoSaveEnabled: () => {},
      print: () => {},
      printPreview: () => {},
      closeWorkbook: () => {},
      quit: () => {},
    },
    openCommandPalette: () => {},
  });

  // View/Developer macro commands are registered separately from `registerDesktopCommands`
  // because they require panel focus wiring + macro-recorder integration in the desktop shell.
  registerRibbonMacroCommands({
    commandRegistry,
    handlers: {
      openPanel: () => {},
      focusScriptEditorPanel: () => {},
      focusVbaMigratePanel: () => {},
      setPendingMacrosPanelFocus: () => {},
      startMacroRecorder: () => {},
      stopMacroRecorder: () => {},
      isTauri: () => false,
    },
  });

  // Power Query / Data → Queries & Connections ribbon commands are registered outside of
  // `registerDesktopCommands` because they depend on PowerQueryService wiring.
  registerDataQueriesCommands({
    commandRegistry,
    layoutController,
    getPowerQueryService: () => null,
    showToast: () => {},
    notify: async () => {},
    refreshRibbonUiState: () => {},
    focusAfterExecute: () => {},
  });

  registerFormatPainterCommand({
    commandRegistry,
    isArmed: () => false,
    arm: () => {},
    disarm: () => {},
  });
}

describe("Ribbon command wiring ↔ CommandRegistry disabling", () => {
  it("does not auto-disable ribbon ids that are explicitly wired in main.ts", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);
    const schemaIdSet = new Set(schemaCommandIds);

    const implementedIds = extractImplementedCommandIdsFromMainTs(schemaIdSet);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(implementedIds).toContain("file.save.save");
    expect(implementedIds).toContain("home.alignment.mergeCenter.mergeCenter");

    const commandRegistry = new CommandRegistry();
    registerCommandsForRibbonDisablingTest(commandRegistry);

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema: defaultRibbonSchema });

    const disabledImplemented = implementedIds.filter((id) => disabledById[id]);
    expect(
      disabledImplemented,
      `Found ribbon ids that are wired in main.ts but disabled by the CommandRegistry baseline:\n${disabledImplemented
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);
  });

  it("ensures enabled ribbon ids that are not registered as commands are referenced in main.ts", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);
    const schemaIdSet = new Set(schemaCommandIds);
    const dropdownTriggerIds = collectRibbonDropdownTriggerIds(defaultRibbonSchema);

    const commandRegistry = new CommandRegistry();
    registerCommandsForRibbonDisablingTest(commandRegistry);

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema: defaultRibbonSchema });

    const enabledButUnregistered = schemaCommandIds
      .filter((id) => !dropdownTriggerIds.has(id))
      .filter((id) => commandRegistry.getCommand(id) == null)
      .filter((id) => !disabledById[id]);

    // Guard: we should always have at least one exempt/non-command ribbon id (e.g. File tab wiring).
    expect(enabledButUnregistered.length).toBeGreaterThan(0);

    const implementedIds = extractImplementedCommandIdsFromMainTs(schemaIdSet);
    const implementedSet = new Set(implementedIds);

    const missing = enabledButUnregistered.filter((id) => !implementedSet.has(id)).sort();
    expect(
      missing,
      [
        "Found ribbon ids that would be enabled but are not registered as commands and are not handled in main.ts.",
        "These ids should either be registered as builtin commands, explicitly handled in the desktop shell,",
        "or removed from the CommandRegistry exemption list so they are disabled by default.",
        "",
        ...missing.map((id) => `- ${id}`),
      ].join("\n"),
    ).toEqual([]);
  });
});
