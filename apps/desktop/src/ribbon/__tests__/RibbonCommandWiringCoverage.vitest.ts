import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../__tests__/sourceTextUtils";

import { CommandRegistry } from "../../extensions/commandRegistry";
import { createDefaultLayout, openPanel, closePanel } from "../../layout/layoutState";
import { panelRegistry } from "../../panels/panelRegistry";
import { registerDesktopCommands } from "../../commands/registerDesktopCommands";

import { computeRibbonDisabledByIdFromCommandRegistry } from "../ribbonCommandRegistryDisabling";
import { handledRibbonCommandIds } from "../ribbonCommandRouter";
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

function collectRibbonSchemaDisabledIds(schema: RibbonSchema): Set<string> {
  // An id is considered "schema-disabled" only if every occurrence in the schema is disabled.
  // This avoids false positives when the same id appears in multiple places with mixed `disabled` flags.
  const disabledCandidates = new Set<string>();
  const enabledIds = new Set<string>();

  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        if (button.disabled) disabledCandidates.add(button.id);
        else enabledIds.add(button.id);

        for (const item of button.menuItems ?? []) {
          if (item.disabled) disabledCandidates.add(item.id);
          else enabledIds.add(item.id);
        }
      }
    }
  }

  return new Set([...disabledCandidates].filter((id) => !enabledIds.has(id)));
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
      { label: "ribbon/ribbonCommandRouter.ts", path: fileURLToPath(new URL("../ribbonCommandRouter.ts", import.meta.url)) },
      { label: "ribbon/commandHandlers.ts", path: fileURLToPath(new URL("../commandHandlers.ts", import.meta.url)) },
    ];

    for (const { label, path } of sources) {
      const source = stripComments(readFileSync(path, "utf8"));

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

function extractImplementedCommandIdsFromDesktopRibbonFallbackHandlers(schemaCommandIds: Set<string>): string[] {
  // The ribbon router provides an explicit, testable list of ids that have real implementations
  // (either via CommandRegistry wiring or ribbon-only handlers). Use that authoritative allowlist
  // so this test doesn't rely on brittle source parsing.
  return Array.from(handledRibbonCommandIds)
    .filter((id) => schemaCommandIds.has(id))
    .sort();
}

function extractImplementedCommandIdsFromRibbonCommandHandlersTs(schemaCommandIds: Set<string>): Set<string> {
  const handlerPath = fileURLToPath(new URL("../commandHandlers.ts", import.meta.url));
  const source = stripComments(readFileSync(handlerPath, "utf8"));
  const ids = new Set<string>();
  for (const match of source.matchAll(/case\s+["']([^"']+)["']/g)) {
    const id = match[1]!;
    if (schemaCommandIds.has(id)) ids.add(id);
  }
  return ids;
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
    // Some registered commands reference split APIs; keep them present so registrations
    // don't drift when invoked from unit tests.
    setSplitDirection(direction: string, ratio?: number) {
      this.layout = {
        ...this.layout,
        splitView: {
          ...(this.layout as any).splitView,
          direction,
          ratio: typeof ratio === "number" ? ratio : (this.layout as any)?.splitView?.ratio ?? 0.5,
        },
      };
    },
  } as any;

  registerDesktopCommands({
    commandRegistry,
    app: {} as any,
    layoutController,
    formatPainter: {
      isArmed: () => false,
      arm: () => {},
      disarm: () => {},
    },
    ribbonMacroHandlers: {
      openPanel: () => {},
      focusScriptEditorPanel: () => {},
      focusVbaMigratePanel: () => {},
      setPendingMacrosPanelFocus: () => {},
      startMacroRecorder: () => {},
      stopMacroRecorder: () => {},
      isTauri: () => false,
    },
    dataQueriesHandlers: {
      getPowerQueryService: () => null,
      showToast: () => {},
      notify: async () => {},
      focusAfterExecute: () => {},
    },
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
    sheetStructureHandlers: {
      insertSheet: () => {},
      deleteActiveSheet: () => {},
      openOrganizeSheets: () => {},
    },
  });
}

describe("Ribbon command wiring ↔ CommandRegistry disabling", () => {
  it("does not auto-disable ribbon ids that are explicitly wired in desktop ribbon fallback handlers", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);
    const schemaIdSet = new Set(schemaCommandIds);
    const schemaDisabledIds = collectRibbonSchemaDisabledIds(defaultRibbonSchema);

    const implementedIds = extractImplementedCommandIdsFromDesktopRibbonFallbackHandlers(schemaIdSet);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(implementedIds).toContain("home.cells.format");
    expect(implementedIds).toContain("home.alignment.mergeCenter.mergeCenter");

    const commandRegistry = new CommandRegistry();
    registerCommandsForRibbonDisablingTest(commandRegistry);

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema: defaultRibbonSchema });

    const disabledImplemented = implementedIds.filter((id) => !schemaDisabledIds.has(id) && disabledById[id]);
    expect(
      disabledImplemented,
      `Found ribbon ids that are wired in desktop ribbon fallback handlers but disabled by the CommandRegistry baseline:\n${disabledImplemented
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);
  });

  it("ensures all enabled ribbon ids are registered CommandRegistry commands (no baseline exemptions needed)", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);
    const dropdownTriggerIds = collectRibbonDropdownTriggerIds(defaultRibbonSchema);
    const schemaDisabledIds = collectRibbonSchemaDisabledIds(defaultRibbonSchema);

    const commandRegistry = new CommandRegistry();
    registerCommandsForRibbonDisablingTest(commandRegistry);

    const disabledById = computeRibbonDisabledByIdFromCommandRegistry(commandRegistry, { schema: defaultRibbonSchema });

    const enabledButUnregistered = schemaCommandIds
      .filter((id) => !dropdownTriggerIds.has(id))
      .filter((id) => !schemaDisabledIds.has(id))
      .filter((id) => commandRegistry.getCommand(id) == null)
      .filter((id) => !disabledById[id]);

    expect(
      enabledButUnregistered,
      [
        "Found ribbon ids that would be enabled but are not registered as CommandRegistry commands.",
        "These ids should be registered as builtin commands (preferred) or disabled by default.",
        "",
        ...enabledButUnregistered.map((id) => `- ${id}`),
      ].join("\n"),
    ).toEqual([]);
  });
});
