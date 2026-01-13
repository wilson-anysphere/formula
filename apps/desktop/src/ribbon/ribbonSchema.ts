import type { RibbonIconId } from "./icons/index.js";

import { fileTab } from "./schema/fileTab.js";
import { homeTab } from "./schema/homeTab.js";
import { insertTab } from "./schema/insertTab.js";
import { pageLayoutTab } from "./schema/pageLayoutTab.js";
import { formulasTab } from "./schema/formulasTab.js";
import { dataTab } from "./schema/dataTab.js";
import { reviewTab } from "./schema/reviewTab.js";
import { viewTab } from "./schema/viewTab.js";
import { developerTab } from "./schema/developerTab.js";
import { helpTab } from "./schema/helpTab.js";

export type RibbonButtonKind = "button" | "toggle" | "dropdown";
export type RibbonButtonSize = "large" | "small" | "icon";

export interface RibbonMenuItemDefinition {
  /**
   * Stable command identifier (used for wiring actions).
   */
  id: string;
  label: string;
  ariaLabel: string;
  /**
   * Stable icon identifier.
   */
  iconId?: RibbonIconId;
  /**
   * Optional E2E hook.
   */
  testId?: string;
  disabled?: boolean;
}

export interface RibbonButtonDefinition {
  /**
   * Stable command identifier (used for wiring actions).
   *
   * Convention: prefer canonical CommandRegistry ids when available (e.g. `clipboard.copy`),
   * otherwise use `{tab}.{group}.{command}` (e.g. `home.font.bold`).
   */
  id: string;
  label: string;
  ariaLabel: string;
  /**
   * Stable icon identifier.
   */
  iconId?: RibbonIconId;
  kind?: RibbonButtonKind;
  size?: RibbonButtonSize;
  /**
   * Optional dropdown menu items. When provided for a `kind: "dropdown"` button,
   * the ribbon will render a menu instead of invoking the command directly.
   */
  menuItems?: RibbonMenuItemDefinition[];
  /**
   * Optional E2E hook.
   */
  testId?: string;
  /**
   * Initial pressed state for toggle buttons (purely UI; can be replaced with
   * app-driven state later).
   */
  defaultPressed?: boolean;
  disabled?: boolean;
}

export interface RibbonGroupDefinition {
  id: string;
  label: string;
  buttons: RibbonButtonDefinition[];
}

export interface RibbonTabDefinition {
  id: string;
  label: string;
  groups: RibbonGroupDefinition[];
  /**
   * File tab is typically styled as a primary pill and may later open a
   * backstage view.
   */
  isFile?: boolean;
}

export interface RibbonSchema {
  tabs: RibbonTabDefinition[];
}

export interface RibbonFileActions {
  newWorkbook?: () => void;
  openWorkbook?: () => void;
  saveWorkbook?: () => void;
  saveWorkbookAs?: () => void;
  /**
   * Toggle the desktop AutoSave feature.
   *
   * This intentionally takes the *next* state (rather than "toggle") so callers
   * can keep UI and persisted state in sync even when the user cancels a required
   * Save As flow.
   */
  toggleAutoSave?: (enabled: boolean) => void;
  versionHistory?: () => void;
  branchManager?: () => void;
  print?: () => void;
  printPreview?: () => void;
  pageSetup?: () => void;
  closeWindow?: () => void;
  quit?: () => void;
}

export interface RibbonActions {
  /**
   * Called when a non-toggle ribbon control is activated.
   *
   * This is emitted for:
   * - Regular buttons (`kind: "button"`).
   * - Dropdown menu items (synthetic buttons generated from `menuItems`).
   * - Dropdown buttons without `menuItems` (treated as a normal command button).
   *
   * Toggle buttons (`kind: "toggle"`) do **not** invoke `onCommand`; they are
   * handled exclusively via `onToggle`.
   */
  onCommand?: (commandId: string) => void;
  /**
   * Called when a toggle button changes state.
   *
   * `pressed` is the *next* pressed state after the user activates the toggle.
   */
  onToggle?: (commandId: string, pressed: boolean) => void;
  /**
   * Called when a tab is selected.
   */
  onTabChange?: (tabId: string) => void;
  /**
   * Optional File tab / backstage actions.
   *
   * The File tab is treated specially (Excel-style "backstage" view) and is
   * wired directly to app-level file operations in `apps/desktop/src/main.ts`.
   */
  fileActions?: RibbonFileActions;
}

export const defaultRibbonSchema: RibbonSchema = {
  tabs: [
    fileTab,
    homeTab,
    insertTab,
    pageLayoutTab,
    formulasTab,
    dataTab,
    reviewTab,
    viewTab,
    developerTab,
    helpTab,
  ],
};

