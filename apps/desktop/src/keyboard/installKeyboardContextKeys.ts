import type { ContextKeyService } from "../extensions/contextKeys.js";
import type { SpreadsheetApp } from "../app/spreadsheetApp";

/**
 * Centralized keyboard/focus-related context keys for `when` clauses.
 *
 * These keys are intended to help route shortcuts through `KeybindingService`
 * without needing ad-hoc DOM target checks scattered throughout the codebase.
 *
 * Keys maintained:
 * - `focus.inTextInput`: `document.activeElement` is INPUT/TEXTAREA/contentEditable
 * - `focus.inFormulaBar`: `document.activeElement` is within the formula bar root
 * - `focus.inSheetTabs`: `document.activeElement` is within the sheet tab strip (tablist)
 * - `focus.inSheetTabRename`: `document.activeElement` is within sheet tabs root *and* is a text input
 * - `focus.inGrid`: `document.activeElement` is within either the primary or secondary grid root
 * - `focus.inSelectionPane`: `document.activeElement` is within the Selection Pane panel (drawing listbox)
 * - `spreadsheet.isEditing`: SpreadsheetApp is editing *or* split-view secondary editor is editing
 * - `spreadsheet.formulaBarEditing`: `app.isFormulaBarEditing()`
 * - `spreadsheet.formulaBarFormulaEditing`: `app.isFormulaBarFormulaEditing()`
 * - `workbench.commandPaletteOpen`: optional hook (defaults to `false`)
 */

export const KeyboardContextKeyIds = {
  focusInTextInput: "focus.inTextInput",
  focusInFormulaBar: "focus.inFormulaBar",
  focusInSheetTabs: "focus.inSheetTabs",
  focusInSheetTabRename: "focus.inSheetTabRename",
  focusInGrid: "focus.inGrid",
  focusInSelectionPane: "focus.inSelectionPane",
  spreadsheetIsEditing: "spreadsheet.isEditing",
  spreadsheetFormulaBarEditing: "spreadsheet.formulaBarEditing",
  spreadsheetFormulaBarFormulaEditing: "spreadsheet.formulaBarFormulaEditing",
  workbenchCommandPaletteOpen: "workbench.commandPaletteOpen",
} as const;

export type KeyboardContextKeyId = (typeof KeyboardContextKeyIds)[keyof typeof KeyboardContextKeyIds];

type SpreadsheetAppKeyboardContext = Pick<
  SpreadsheetApp,
  | "isEditing"
  | "isFormulaBarEditing"
  | "isFormulaBarFormulaEditing"
  | "onEditStateChange"
  | "onFormulaBarOverlayChange"
>;

export type KeyboardContextKeysParams = {
  contextKeys: ContextKeyService;
  app: SpreadsheetAppKeyboardContext;
  formulaBarRoot: HTMLElement;
  sheetTabsRoot: HTMLElement;
  gridRoot: HTMLElement;
  gridSecondaryRoot?: HTMLElement | null;
  // Optional hooks:
  isCommandPaletteOpen?: () => boolean;
  isSplitViewSecondaryEditing?: () => boolean;
};

export type KeyboardContextKeysDisposer = (() => void) & {
  /**
   * Force a recompute. Useful for state that isn't covered by focus/app subscriptions
   * (e.g. split-view secondary editor edit state).
   */
  recompute: () => void;
};

function isTextInputLike(target: HTMLElement | null): boolean {
  if (!target) return false;
  const tag = target.tagName;
  // `HTMLElement.isContentEditable` is a boolean in browsers, but some DOM shims (jsdom)
  // may not implement it consistently. Coerce to boolean so downstream context keys
  // are always strict booleans.
  return tag === "INPUT" || tag === "TEXTAREA" || Boolean((target as any).isContentEditable);
}

function safeContains(root: HTMLElement, el: HTMLElement | null): boolean {
  if (!el) return false;
  try {
    return root.contains(el);
  } catch {
    return false;
  }
}

export function installKeyboardContextKeys(params: KeyboardContextKeysParams): KeyboardContextKeysDisposer {
  const { contextKeys, app, formulaBarRoot, sheetTabsRoot, gridRoot, gridSecondaryRoot, isCommandPaletteOpen, isSplitViewSecondaryEditing } =
    params;

  let disposed = false;

  const recompute = (): void => {
    if (disposed) return;

    const active = typeof document !== "undefined" ? (document.activeElement as HTMLElement | null) : null;
    const inTextInput = isTextInputLike(active);

    const inFormulaBar = safeContains(formulaBarRoot, active);
    const sheetTabStripRoot = (() => {
      try {
        // `#sheet-tabs` contains other buttons (add/overflow) that are not part of the tablist.
        // Prefer scoping to the actual tab strip when available so context keys can discriminate.
        return sheetTabsRoot.querySelector<HTMLElement>(".sheet-tabs") ?? sheetTabsRoot;
      } catch {
        return sheetTabsRoot;
      }
    })();
    const inSheetTabs = safeContains(sheetTabStripRoot, active);
    const inSheetTabsRoot = safeContains(sheetTabsRoot, active);
    const inSheetTabRename = inSheetTabsRoot && inTextInput;
    const inGrid = safeContains(gridRoot, active) || (gridSecondaryRoot ? safeContains(gridSecondaryRoot, active) : false);
    const inSelectionPane = (() => {
      if (!active) return false;
      try {
        return typeof active.closest === "function" && Boolean(active.closest(".selection-pane"));
      } catch {
        return false;
      }
    })();

    const secondaryEditing = isSplitViewSecondaryEditing?.() === true;
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    const isEditing = Boolean(app.isEditing() || secondaryEditing || globalEditing === true);

    contextKeys.batch({
      [KeyboardContextKeyIds.focusInTextInput]: inTextInput,
      [KeyboardContextKeyIds.focusInFormulaBar]: inFormulaBar,
      [KeyboardContextKeyIds.focusInSheetTabs]: inSheetTabs,
      [KeyboardContextKeyIds.focusInSheetTabRename]: inSheetTabRename,
      [KeyboardContextKeyIds.focusInGrid]: inGrid,
      [KeyboardContextKeyIds.focusInSelectionPane]: inSelectionPane,
      [KeyboardContextKeyIds.spreadsheetIsEditing]: isEditing,
      [KeyboardContextKeyIds.spreadsheetFormulaBarEditing]: Boolean(app.isFormulaBarEditing()),
      [KeyboardContextKeyIds.spreadsheetFormulaBarFormulaEditing]: Boolean(app.isFormulaBarFormulaEditing()),
      [KeyboardContextKeyIds.workbenchCommandPaletteOpen]: Boolean(isCommandPaletteOpen?.() ?? false),
    });
  };

  const onFocusEvent = (): void => recompute();

  // Focus changes drive most "should shortcuts run?" decisions.
  // Use capture so we see focus changes before other handlers potentially stop propagation.
  document.addEventListener("focusin", onFocusEvent, true);
  document.addEventListener("focusout", onFocusEvent, true);

  // Spreadsheet editing state (cell editor, formula bar, inline edit controller).
  const unsubscribeEditState = app.onEditStateChange(() => recompute());

  // Formula-bar overlay changes can update "formula editing mode" while the user types (eg `=`).
  const unsubscribeFormulaOverlay = app.onFormulaBarOverlayChange(() => recompute());

  // Ensure initial keys are available immediately after install.
  recompute();

  const dispose = (() => {
    if (disposed) return;
    disposed = true;
    document.removeEventListener("focusin", onFocusEvent, true);
    document.removeEventListener("focusout", onFocusEvent, true);
    unsubscribeEditState();
    unsubscribeFormulaOverlay();
  }) as KeyboardContextKeysDisposer;

  dispose.recompute = recompute;

  return dispose;
}
