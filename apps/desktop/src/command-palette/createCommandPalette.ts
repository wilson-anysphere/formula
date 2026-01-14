import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import type { ContextKeyService } from "../extensions/contextKeys.js";
import { isSpreadsheetEditingCommandBlockedError } from "../commands/spreadsheetEditingCommandBlockedError.js";

import { t, tWithVars } from "../i18n/index.js";
import { markKeybindingBarrier } from "../keybindingBarrier.js";
import { evaluateWhenClause } from "../extensions/whenClause.js";

import { debounce } from "./debounce.js";
import {
  compileFuzzyQuery,
  fuzzyMatchCommandPrepared,
  prepareCommandForFuzzy,
  type MatchRange,
  type PreparedCommandForFuzzy,
} from "./fuzzy.js";
import { searchFunctionResults, type CommandPaletteFunctionResult } from "./commandPaletteSearch.js";
import { getRecentCommandIdsForDisplay, type StorageLike } from "./recents.js";
import { installCommandPaletteRecentsTracking } from "./installCommandPaletteRecentsTracking.js";
import { searchShortcutCommands } from "./shortcutSearch.js";
import { isFunctionSignatureCatalogReady, preloadFunctionSignatureCatalog } from "../formula-bar/highlight/functionSignatures.js";
import { formatA1Range, parseGoTo, type GoToParseResult, type GoToWorkbookLookup } from "../../../../packages/search/index.js";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";

type RenderableCommand = PreparedCommandForFuzzy<CommandContribution> & {
  score: number;
  titleRanges: MatchRange[];
  /**
   * Optional shortcut override used by shortcut-search mode.
   *
   * When a command has multiple keybindings, shortcut search may prefer displaying the binding
   * that matched the query (instead of always showing the primary binding).
   */
  shortcut?: string;
};

type CommandGroup = {
  label: string;
  commands: RenderableCommand[];
};

type GoToSuggestion = {
  kind: "goTo";
  label: string;
  resolved: string;
  parsed: GoToParseResult;
};

type RenderableItem = { kind: "command"; command: RenderableCommand } | GoToSuggestion | CommandPaletteFunctionResult;

export type CreateCommandPaletteOptions = {
  commandRegistry: CommandRegistry;
  contextKeys: ContextKeyService;
  /**
   * A `commandId -> [displayKeybinding...]` index. The palette renders the first
   * entry as the primary shortcut hint.
   */
  keybindingIndex: Map<string, string | readonly string[]>;
  ensureExtensionsLoaded: () => Promise<void>;
  onCloseFocus: () => void;
  placeholder?: string;
  goTo?:
    | {
        workbook: GoToWorkbookLookup;
        getCurrentSheetName: () => string;
        onGoTo: (parsed: GoToParseResult) => void;
      }
    | null;
  /**
   * How long to wait after opening before we kick off extension loading.
   * Defaults to 600ms to avoid paying the extension-worker cost for "quick" palette usage
   * (notably in e2e tests that just invoke built-in commands).
   */
  extensionLoadDelayMs?: number;
  /**
   * Hard cap the number of command rows rendered (excludes group headers).
   * Defaults to 100 to prevent DOM blowups with large command/function sets.
   */
  maxResults?: number;
  /**
   * Optional per-category cap (applied after global ranking).
   * Defaults to 20 to keep large categories from dominating.
   */
  maxResultsPerGroup?: number;
  /**
   * Debounce input updates to keep typing responsive when rescoring large lists.
   * Defaults to 70ms.
   */
  inputDebounceMs?: number;
  /**
   * Optional handler for selecting a spreadsheet function result from the palette.
   */
  onSelectFunction?: (name: string) => void;
};

export type CommandPaletteHandle = {
  open: () => void;
  close: () => void;
  isOpen: () => boolean;
  dispose: () => void;
};

function groupLabel(category: string | null): string {
  const value = String(category ?? "").trim();
  return value ? value : t("commandPalette.group.other");
}

function sortCommandsAlpha(a: CommandContribution, b: CommandContribution): number {
  return a.title.localeCompare(b.title);
}

function buildGroupsForEmptyQuery(
  allCommands: PreparedCommandForFuzzy<CommandContribution>[],
  recentIds: string[],
  limits: { maxResults: number; maxResultsPerGroup: number },
): CommandGroup[] {
  const byId = new Map(allCommands.map((cmd) => [cmd.commandId, cmd]));

  const recents: RenderableCommand[] = [];
  const recentSet = new Set<string>();
  const recentLimit = Math.min(limits.maxResultsPerGroup, limits.maxResults);
  for (const id of recentIds) {
    if (recents.length >= recentLimit) break;
    const cmd = byId.get(id);
    if (!cmd) continue;
    recentSet.add(id);
    recents.push({ ...cmd, score: 0, titleRanges: [] });
  }

  const remaining = allCommands.filter((cmd) => !recentSet.has(cmd.commandId));

  // Empty query can include thousands of commands (builtins + extensions + functions).
  // Avoid sorting enormous arrays per category by keeping only the alphabetically-best
  // `maxResultsPerGroup` commands for each category as we scan.
  const categories = new Map<string, RenderableCommand[]>();
  for (const cmd of remaining) {
    const label = groupLabel(cmd.category);
    const list = categories.get(label) ?? [];

    if (list.length < limits.maxResultsPerGroup) {
      list.push({ ...cmd, score: 0, titleRanges: [] });
      categories.set(label, list);
      continue;
    }

    // Find the alphabetically-last command (worst) in this limited list and replace it
    // if the new command sorts before it.
    let worstIdx = 0;
    for (let i = 1; i < list.length; i += 1) {
      if (list[i]!.title.localeCompare(list[worstIdx]!.title) > 0) worstIdx = i;
    }
    if (cmd.title.localeCompare(list[worstIdx]!.title) < 0) {
      list[worstIdx] = { ...cmd, score: 0, titleRanges: [] };
      categories.set(label, list);
    }
  }

  const groups: CommandGroup[] = [];
  let remainingSlots = Math.max(0, limits.maxResults - recents.length);

  if (recents.length > 0) groups.push({ label: t("commandPalette.group.recent"), commands: recents });

  const sortedCategoryLabels = [...categories.keys()].sort((a, b) => a.localeCompare(b));
  for (const label of sortedCategoryLabels) {
    if (remainingSlots <= 0) break;
    const cmds = categories.get(label)!;
    cmds.sort(sortCommandsAlpha);
    const slice = cmds.slice(0, Math.min(limits.maxResultsPerGroup, remainingSlots));
    if (slice.length === 0) continue;
    groups.push({ label, commands: slice });
    remainingSlots -= slice.length;
  }

  return groups;
}

function buildGroupsForQuery(
  allCommands: PreparedCommandForFuzzy<CommandContribution>[],
  query: string,
  limits: { maxResults: number; maxResultsPerGroup: number },
): CommandGroup[] {
  if (limits.maxResults <= 0) return [];
  const compiled = compileFuzzyQuery(query);

  // Keep only top-N matches to avoid sorting huge arrays. We keep more than we
  // ultimately render so that per-group caps still yield ~maxResults total items.
  const maxRanked = Math.min(allCommands.length, limits.maxResults * 3);
  const top: RenderableCommand[] = [];

  const isBetter = (a: RenderableCommand, b: RenderableCommand): boolean => {
    if (a.score !== b.score) return a.score > b.score;
    return a.title.localeCompare(b.title) < 0;
  };

  const worstIndex = (): number => {
    let worst = 0;
    for (let i = 1; i < top.length; i += 1) {
      // If the current `worst` is better than `top[i]`, then `top[i]` is the new worst.
      if (isBetter(top[worst]!, top[i]!)) worst = i;
    }
    return worst;
  };

  for (const cmd of allCommands) {
    const match = fuzzyMatchCommandPrepared(compiled, cmd);
    if (!match) continue;

    const candidate: RenderableCommand = { ...cmd, score: match.score, titleRanges: match.titleRanges };
    if (top.length < maxRanked) {
      top.push(candidate);
      continue;
    }

    const idx = worstIndex();
    if (isBetter(candidate, top[idx]!)) {
      top[idx] = candidate;
    }
  }

  top.sort((a, b) => {
    if (a.score !== b.score) return b.score - a.score;
    return a.title.localeCompare(b.title);
  });

  const groupsByLabel = new Map<string, RenderableCommand[]>();
  const groupOrder: string[] = [];

  let remainingSlots = limits.maxResults;
  for (const cmd of top) {
    if (remainingSlots <= 0) break;
    const label = groupLabel(cmd.category);
    if (!groupsByLabel.has(label)) groupOrder.push(label);
    const list = groupsByLabel.get(label) ?? [];
    if (list.length < limits.maxResultsPerGroup) {
      list.push(cmd);
      groupsByLabel.set(label, list);
      remainingSlots -= 1;
    }
  }

  return groupOrder.map((label) => ({ label, commands: groupsByLabel.get(label) ?? [] }));
}

function buildGroupsForShortcutMode(
  matches: Array<PreparedCommandForFuzzy<CommandContribution> & { shortcut: string }>,
  limits: { maxResults: number; maxResultsPerGroup: number },
): CommandGroup[] {
  if (limits.maxResults <= 0) return [];
  const groupsByLabel = new Map<string, RenderableCommand[]>();
  const order: string[] = [];

  let remainingSlots = limits.maxResults;
  for (const cmd of matches) {
    if (remainingSlots <= 0) break;
    const label = groupLabel(cmd.category);
    if (!groupsByLabel.has(label)) order.push(label);
    const list = groupsByLabel.get(label) ?? [];
    if (list.length < limits.maxResultsPerGroup) {
      list.push({ ...cmd, score: 0, titleRanges: [] });
      groupsByLabel.set(label, list);
      remainingSlots -= 1;
    }
  }

  return order.map((label) => ({ label, commands: groupsByLabel.get(label) ?? [] }));
}

function renderHighlightedText(text: string, ranges: MatchRange[]): DocumentFragment {
  const fragment = document.createDocumentFragment();
  if (!ranges.length) {
    fragment.appendChild(document.createTextNode(text));
    return fragment;
  }

  let pos = 0;
  for (const range of ranges) {
    if (range.start > pos) fragment.appendChild(document.createTextNode(text.slice(pos, range.start)));
    const span = document.createElement("span");
    span.className = "command-palette__highlight";
    span.textContent = text.slice(range.start, range.end);
    fragment.appendChild(span);
    pos = range.end;
  }

  if (pos < text.length) fragment.appendChild(document.createTextNode(text.slice(pos)));
  return fragment;
}

export function createCommandPalette(options: CreateCommandPaletteOptions): CommandPaletteHandle {
  const {
    commandRegistry,
    contextKeys,
    keybindingIndex,
    ensureExtensionsLoaded,
    onCloseFocus,
    placeholder = t("commandPalette.placeholder"),
    goTo = null,
    extensionLoadDelayMs = 600,
    maxResults = 100,
    maxResultsPerGroup = 20,
    inputDebounceMs = 70,
    onSelectFunction,
  } = options;

  const COMMAND_PALETTE_OPEN_CONTEXT_KEY = "workbench.commandPaletteOpen";
  // Ensure the key is always defined so `workbench.commandPaletteOpen == false` works
  // (and so `!workbench.commandPaletteOpen` behaves consistently across runtimes).
  contextKeys.set(COMMAND_PALETTE_OPEN_CONTEXT_KEY, false);

  const overlay = document.createElement("div");
  overlay.className = "command-palette-overlay";
  overlay.dataset.keybindingBarrier = "true";
  overlay.hidden = true;
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");
  overlay.setAttribute("aria-label", t("commandPalette.aria.label"));
  markKeybindingBarrier(overlay);

  const palette = document.createElement("div");
  palette.className = "command-palette";
  palette.dataset.testid = "command-palette";

  const input = document.createElement("input");
  input.className = "command-palette__input";
  input.dataset.testid = "command-palette-input";
  input.placeholder = placeholder;
  // Treat the palette input as a combobox controlling the listbox so assistive
  // tech can follow `aria-activedescendant` updates while focus stays in the input.
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-autocomplete", "list");
  input.setAttribute("aria-expanded", "false");
  input.setAttribute("aria-label", placeholder);

  const hint = document.createElement("div");
  hint.className = "command-palette__hint";
  hint.textContent = t("commandPalette.shortcutSearch.hint");
  hint.hidden = true;

  const list = document.createElement("ul");
  list.className = "command-palette__list";
  list.dataset.testid = "command-palette-list";
  list.id = "command-palette-listbox";
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", t("commandPalette.aria.commandsList"));
  // Ensure there's always a second tabbable target for the focus trap.
  list.tabIndex = 0;
  input.setAttribute("aria-controls", list.id);
  input.setAttribute("aria-haspopup", "listbox");

  palette.appendChild(input);
  palette.appendChild(hint);
  palette.appendChild(list);
  overlay.appendChild(palette);
  document.body.appendChild(overlay);

  const storage: StorageLike = (() => {
    try {
      return localStorage;
    } catch {
      return {
        getItem: () => null,
        setItem: () => {},
      };
    }
  })();

  const limits = {
    maxResults: Math.max(0, Math.floor(maxResults)),
    maxResultsPerGroup: Math.max(1, Math.floor(maxResultsPerGroup)),
  };
  const inputDebounce = Math.max(0, Math.floor(inputDebounceMs));

  let isOpen = false;
  let query = "";
  let selectedIndex = 0;
  let visibleItems: RenderableItem[] = [];
  let visibleItemEls: HTMLLIElement[] = [];
  let commandsCacheDirty = true;
  let cachedCommands: PreparedCommandForFuzzy<CommandContribution>[] = [];
  let chunkSearchController: AbortController | null = null;
  let extensionLoadTimer: number | null = null;
  let lastFocusedElement: HTMLElement | null = null;

  const handleDocumentFocusIn = (e: FocusEvent): void => {
    if (!isOpen) return;
    const target = e.target as Node | null;
    if (!target) return;
    if (overlay.contains(target)) return;
    // If focus escapes the modal dialog, bring it back to the input.
    input.focus();
  };

  const handleOverlayKeyDown = (e: KeyboardEvent): void => {
    if (!isOpen) return;
    if (e.key !== "Tab") return;

    // Minimal focus trap: cycle focus between the input and listbox.
    const focusable = [input, list].filter((el) => !el.hasAttribute("disabled"));
    if (focusable.length === 0) return;

    e.preventDefault();

    const active = document.activeElement as HTMLElement | null;
    const currentIndex = active ? focusable.indexOf(active as (typeof focusable)[number]) : -1;
    const delta = e.shiftKey ? -1 : 1;
    const nextIndex =
      currentIndex === -1 ? (e.shiftKey ? focusable.length - 1 : 0) : (currentIndex + delta + focusable.length) % focusable.length;
    focusable[nextIndex]!.focus();
  };
  overlay.addEventListener("keydown", handleOverlayKeyDown);

  const executeCommand = (commandId: string): void => {
    void commandRegistry.executeCommand(commandId).catch((err) => {
      if (isSpreadsheetEditingCommandBlockedError(err)) return;
      console.error(`Command failed (${commandId}):`, err);
    });
  };

  const abortChunkedSearch = (): void => {
    chunkSearchController?.abort();
    chunkSearchController = null;
  };

  const setActiveDescendant = (id: string | null): void => {
    if (!id) {
      list.removeAttribute("aria-activedescendant");
      input.removeAttribute("aria-activedescendant");
      return;
    }
    list.setAttribute("aria-activedescendant", id);
    input.setAttribute("aria-activedescendant", id);
  };

  function close(): void {
    // Keep the context key authoritative even if `close()` is called redundantly.
    contextKeys.set(COMMAND_PALETTE_OPEN_CONTEXT_KEY, false);
    if (!isOpen) return;
    isOpen = false;
    debouncedRender.cancel();
    abortChunkedSearch();
    overlay.hidden = true;
    document.removeEventListener("focusin", handleDocumentFocusIn);
    input.setAttribute("aria-expanded", "false");
    setActiveDescendant(null);
    query = "";
    selectedIndex = 0;
    visibleItems = [];
    visibleItemEls = [];
    if (extensionLoadTimer != null) {
      window.clearTimeout(extensionLoadTimer);
      extensionLoadTimer = null;
    }

    // Best-effort restore: return focus to whatever triggered opening the palette.
    const restoreTarget = lastFocusedElement;
    lastFocusedElement = null;
    if (restoreTarget && restoreTarget.isConnected && typeof restoreTarget.focus === "function") {
      try {
        restoreTarget.focus();
        if (document.activeElement === restoreTarget) return;
      } catch {
        // ignore
      }
    }

    onCloseFocus();
  }

  function open(): void {
    // Keep the context key authoritative even if `open()` is called redundantly.
    contextKeys.set(COMMAND_PALETTE_OPEN_CONTEXT_KEY, true);
    if (isOpen) {
      input.focus();
      input.select();
      return;
    }

    debouncedRender.cancel();
    abortChunkedSearch();

    const active = document.activeElement;
    lastFocusedElement =
      active instanceof HTMLElement && active !== document.body && !overlay.contains(active) ? active : null;
    query = "";
    selectedIndex = 0;
    input.value = "";
    overlay.hidden = false;
    isOpen = true;
    input.setAttribute("aria-expanded", "true");
    document.addEventListener("focusin", handleDocumentFocusIn);

    // Best-effort: warm the function signature catalog in the background so function results can
    // show signatures/summaries shortly after the palette opens.
    if (!isFunctionSignatureCatalogReady()) {
      void preloadFunctionSignatureCatalog()
        .then(() => {
          if (!isOpen) return;
          if (!isFunctionSignatureCatalogReady()) return;
          if (!query.trim()) return;
          // Re-render to populate signatures for catalog-only functions.
          renderResults("async");
        })
        .catch(() => {
          // Best-effort: ignore catalog prefetch failures.
        });
    }

    renderResults("sync");

    input.focus();
    input.select();

    // Best-effort: load extensions in the background so contributed commands appear,
    // but defer long enough that quick palette usage doesn't pay the cost.
    if (extensionLoadTimer != null) window.clearTimeout(extensionLoadTimer);
    extensionLoadTimer = window.setTimeout(() => {
      extensionLoadTimer = null;
      void ensureExtensionsLoaded().catch(() => {
        // ignore
      });
    }, extensionLoadDelayMs);
  }

  function ensureCommandsCache(): void {
    if (!commandsCacheDirty) return;
    const lookup = contextKeys.asLookup();
    cachedCommands = commandRegistry
      .listCommands()
      // The command palette owns opening itself; avoid showing a no-op entry.
      .filter((cmd) => cmd.commandId !== "workbench.showCommandPalette")
      // Hide commands whose context key expression is not satisfied (e.g. role/permission gated).
      .filter((cmd) => evaluateWhenClause(cmd.when, lookup))
      .map((cmd) => prepareCommandForFuzzy(cmd));
    commandsCacheDirty = false;
  }

  function getRecentsForDisplay(allCommands: Array<{ commandId: string }>): string[] {
    return getRecentCommandIdsForDisplay(
      storage,
      allCommands.map((cmd) => cmd.commandId),
    );
  }

  // Cache command row DOM nodes by id so we don't create new nodes for the same
  // results across renders. Keep this bounded so long-running sessions can't
  // accumulate thousands of detached DOM nodes in memory.
  const rowCacheMax = Math.max(300, limits.maxResults * 5);

  const commandRowCache = new Map<
    string,
    {
      li: HTMLLIElement;
      icon: HTMLDivElement;
      label: HTMLDivElement;
      description: HTMLDivElement;
      right: HTMLDivElement;
      shortcutPill: HTMLSpanElement;
    }
  >();

  const CHUNK_SEARCH_MIN_COMMANDS = 5_000;
  const CHUNK_SEARCH_MIN_QUERY_LEN = 4;
  const CHUNK_SEARCH_CHUNK_SIZE = 500;

  function renderGroups(groups: CommandGroup[], emptyText: string): void {
    const trimmed = query.trim();
    const shortcutMode = trimmed.startsWith("/");

    const goToSuggestion: GoToSuggestion | null =
      !shortcutMode && trimmed !== "" && goTo
        ? (() => {
            try {
              const parsed = parseGoTo(trimmed, {
                workbook: goTo.workbook,
                currentSheetName: goTo.getCurrentSheetName(),
              });
               return {
                 kind: "goTo" as const,
                 label: tWithVars("commandPalette.goToSuggestion", { query: trimmed }),
                 resolved: `${formatSheetNameForA1(parsed.sheetName)}!${formatA1Range(parsed.range)}`,
                 parsed,
               };
             } catch {
               return null;
             }
          })()
        : null;

    const functionResults: CommandPaletteFunctionResult[] =
      onSelectFunction && !shortcutMode && trimmed !== ""
        ? searchFunctionResults(trimmed, { limit: Math.min(limits.maxResults, limits.maxResultsPerGroup) })
        : [];

    const functionsToShow = functionResults.slice(
      0,
      Math.min(functionResults.length, limits.maxResultsPerGroup, limits.maxResults),
    );
    const commandSlots = Math.max(0, limits.maxResults - functionsToShow.length);

    const limitedGroups = (() => {
      if (commandSlots >= limits.maxResults) return groups;
      const out: CommandGroup[] = [];
      let remaining = commandSlots;
      for (const group of groups) {
        if (remaining <= 0) break;
        const slice = group.commands.slice(0, remaining);
        if (slice.length === 0) continue;
        out.push({ label: group.label, commands: slice });
        remaining -= slice.length;
      }
      return out;
    })();

    const commandCount = limitedGroups.reduce((sum, g) => sum + g.commands.length, 0);
    const bestCommand = limitedGroups[0]?.commands[0]?.score ?? Number.NEGATIVE_INFINITY;
    const bestFunction = functionsToShow[0]?.score ?? Number.NEGATIVE_INFINITY;
    const functionsFirst = bestFunction > bestCommand;

    visibleItems = [];
    if (goToSuggestion) visibleItems.push(goToSuggestion);

    if (functionsFirst) visibleItems.push(...functionsToShow);

    for (const group of limitedGroups) {
      for (const cmd of group.commands) {
        visibleItems.push({ kind: "command", command: cmd });
      }
    }

    if (!functionsFirst) visibleItems.push(...functionsToShow);

    if (selectedIndex >= visibleItems.length) selectedIndex = Math.max(0, visibleItems.length - 1);

    list.replaceChildren();
    visibleItemEls = [];
    setActiveDescendant(null);

    if (visibleItems.length === 0) {
      const empty = document.createElement("li");
      empty.className = "command-palette__empty";
      empty.textContent = emptyText;
      empty.setAttribute("role", "option");
      empty.setAttribute("aria-disabled", "true");
      empty.setAttribute("aria-selected", "false");
      list.appendChild(empty);
      return;
    }

    if (goToSuggestion) {
      const li = document.createElement("li");
      li.className = "command-palette__item";
      li.id = "command-palette-option-0";
      li.setAttribute("role", "option");
      li.setAttribute("aria-selected", selectedIndex === 0 ? "true" : "false");

      const icon = document.createElement("div");
      icon.className = "command-palette__item-icon command-palette__item-icon--goto";
      icon.textContent = "↦";

      const main = document.createElement("div");
      main.className = "command-palette__item-main";

      const label = document.createElement("div");
      label.className = "command-palette__item-label";
      label.textContent = goToSuggestion.label;
      main.appendChild(label);

      const description = document.createElement("div");
      description.className = "command-palette__item-description";
      description.textContent = goToSuggestion.resolved;
      main.appendChild(description);

      const right = document.createElement("div");
      right.className = "command-palette__item-right";

      const enterHint = document.createElement("span");
      enterHint.className = "command-palette__shortcut command-palette__selected-hint";
      enterHint.textContent = "↵";
      right.appendChild(enterHint);

      li.appendChild(icon);
      li.appendChild(main);
      li.appendChild(right);

      li.addEventListener("mousedown", (e) => {
        e.preventDefault();
      });
      li.addEventListener("click", () => {
        close();
        goTo?.onGoTo(goToSuggestion.parsed);
      });

      list.appendChild(li);
      visibleItemEls[0] = li;
    }

    const baseOffset = goToSuggestion ? 1 : 0;

    const renderFunctionRows = (startIndex: number): void => {
      if (functionsToShow.length === 0) return;

      const header = document.createElement("li");
      header.className = "command-palette__group";
      header.textContent = t("commandPalette.group.functions");
      header.setAttribute("role", "presentation");
      header.setAttribute("aria-hidden", "true");
      list.appendChild(header);

      for (let i = 0; i < functionsToShow.length; i += 1) {
        const fn = functionsToShow[i]!;
        const globalIndex = startIndex + i;

        const li = document.createElement("li");
        li.className = "command-palette__item";
        li.id = `command-palette-option-${globalIndex}`;
        // Stable hook for Playwright tests (and future UI automation) to select a specific
        // spreadsheet function row without relying on ranking order or DOM structure.
        li.dataset.testid = `command-palette-function-${fn.name}`;
        li.setAttribute("role", "option");
        li.setAttribute("aria-selected", globalIndex === selectedIndex ? "true" : "false");

        const icon = document.createElement("div");
        icon.className = "command-palette__item-icon command-palette__item-icon--function";
        icon.textContent = "Σ";

        const main = document.createElement("div");
        main.className = "command-palette__item-main";

        const label = document.createElement("div");
        label.className = "command-palette__item-label";
        label.appendChild(renderHighlightedText(fn.name, fn.matchRanges));
        main.appendChild(label);

        if (fn.summary) {
          const summary = document.createElement("div");
          summary.className = "command-palette__item-description";
          summary.textContent = fn.summary;
          main.appendChild(summary);
        }

        if (fn.signature) {
          const signature = document.createElement("div");
          signature.className = "command-palette__item-description command-palette__item-description--mono";
          if (fn.signature.startsWith(fn.name)) {
            const fnName = document.createElement("span");
            fnName.className = "command-palette__signature-name";
            fnName.textContent = fn.name;
            signature.appendChild(fnName);
            signature.appendChild(document.createTextNode(fn.signature.slice(fn.name.length)));
          } else {
            signature.textContent = fn.signature;
          }
          main.appendChild(signature);
        }

        const right = document.createElement("div");
        right.className = "command-palette__item-right";

        const enterHint = document.createElement("span");
        enterHint.className = "command-palette__shortcut command-palette__selected-hint";
        enterHint.textContent = "↵";
        right.appendChild(enterHint);

        li.appendChild(icon);
        li.appendChild(main);
        li.appendChild(right);

        li.addEventListener("mousedown", (e) => {
          // Prevent focus leaving the input before we run the command.
          e.preventDefault();
        });
        li.addEventListener("click", () => {
          close();
          onSelectFunction?.(fn.name);
        });

        list.appendChild(li);
        visibleItemEls[globalIndex] = li;
      }
    };

    const renderCommandRows = (startIndex: number): void => {
      let commandOffset = 0;

      const recentLabel = t("commandPalette.group.recent");
      const isEmptyQuery = trimmed === "";
      const hasCommands = limitedGroups.some((group) => group.commands.length > 0);

      // Match the command palette mockup: show RECENT first (when present), then a COMMANDS section
      // for the rest of the command results.
      let startGroupIndex = 0;
      if (isEmptyQuery && limitedGroups[0]?.label === recentLabel) {
        startGroupIndex = 1;
      } else if (hasCommands) {
        const header = document.createElement("li");
        header.className = "command-palette__group";
        header.textContent = t("commandPalette.group.commands");
        header.setAttribute("role", "presentation");
        header.setAttribute("aria-hidden", "true");
        list.appendChild(header);
      }

      for (let groupIndex = 0; groupIndex < limitedGroups.length; groupIndex += 1) {
        const group = limitedGroups[groupIndex]!;

        if (groupIndex === startGroupIndex && hasCommands && startGroupIndex === 1) {
          const header = document.createElement("li");
          header.className = "command-palette__group";
          header.textContent = t("commandPalette.group.commands");
          header.setAttribute("role", "presentation");
          header.setAttribute("aria-hidden", "true");
          list.appendChild(header);
        }

        const header = document.createElement("li");
        header.className = "command-palette__group";
        header.textContent = group.label;
        header.setAttribute("role", "presentation");
        header.setAttribute("aria-hidden", "true");
        list.appendChild(header);

        for (let i = 0; i < group.commands.length; i += 1) {
          const cmd = group.commands[i]!;
          const globalIndex = startIndex + commandOffset + i;

          let cached = commandRowCache.get(cmd.commandId);
          if (!cached) {
            const li = document.createElement("li");
            li.className = "command-palette__item";
            li.setAttribute("role", "option");

            const icon = document.createElement("div");
            icon.className = "command-palette__item-icon command-palette__item-icon--command";
            icon.textContent = "⌘";

            const main = document.createElement("div");
            main.className = "command-palette__item-main";

            const label = document.createElement("div");
            label.className = "command-palette__item-label";

            const description = document.createElement("div");
            description.className = "command-palette__item-description";

            main.appendChild(label);
            main.appendChild(description);

            const right = document.createElement("div");
            right.className = "command-palette__item-right";

            const shortcutPill = document.createElement("span");
            shortcutPill.className = "command-palette__shortcut";
            right.appendChild(shortcutPill);

            li.appendChild(icon);
            li.appendChild(main);
            li.appendChild(right);

            li.addEventListener("mousedown", (e) => {
              // Prevent focus leaving the input before we run the command.
              e.preventDefault();
            });
            li.addEventListener("click", () => {
              close();
              executeCommand(cmd.commandId);
            });

            cached = { li, icon, label, description, right, shortcutPill };
            commandRowCache.set(cmd.commandId, cached);
          } else {
            // Mark as most-recently-used (Map iteration order is insertion order).
            commandRowCache.delete(cmd.commandId);
            commandRowCache.set(cmd.commandId, cached);
          }

          cached.li.id = `command-palette-option-${globalIndex}`;
          cached.li.setAttribute("aria-selected", globalIndex === selectedIndex ? "true" : "false");
          cached.label.replaceChildren(renderHighlightedText(cmd.title, cmd.titleRanges));

          const descriptionText = typeof cmd.description === "string" ? cmd.description.trim() : "";
          cached.description.textContent = descriptionText;
          cached.description.hidden = !descriptionText;

          const shortcutFromShortcutMode = shortcutMode ? cmd.shortcut ?? null : null;
          const kbValue = keybindingIndex.get(cmd.commandId);
          const shortcut =
            shortcutFromShortcutMode ??
            (typeof kbValue === "string" ? kbValue : Array.isArray(kbValue) ? (kbValue[0] ?? null) : null);
          if (shortcut) {
            cached.shortcutPill.textContent = shortcut;
            cached.right.hidden = false;
          } else {
            cached.shortcutPill.textContent = "";
            cached.right.hidden = true;
          }

          list.appendChild(cached.li);
          visibleItemEls[globalIndex] = cached.li;
        }

        commandOffset += group.commands.length;
      }
    };

    if (functionsFirst) {
      renderFunctionRows(baseOffset);
      renderCommandRows(baseOffset + functionsToShow.length);
    } else {
      renderCommandRows(baseOffset);
      renderFunctionRows(baseOffset + commandCount);
    }

    // Keep selection in view after re-render.
    const selectedEl = visibleItemEls[selectedIndex];
    setActiveDescendant(selectedEl?.id ?? null);
    queueMicrotask(() => {
      const el = visibleItemEls[selectedIndex];
      if (el && typeof el.scrollIntoView === "function") {
        el.scrollIntoView({ block: "nearest" });
      }
    });

    // Evict least-recently-used rows to keep the cache bounded.
    while (commandRowCache.size > rowCacheMax) {
      const oldest = commandRowCache.keys().next().value as string | undefined;
      if (!oldest) break;
      commandRowCache.delete(oldest);
    }
  }

  function startChunkedSearch(querySnapshot: string): void {
    ensureCommandsCache();

    abortChunkedSearch();
    const controller = new AbortController();
    chunkSearchController = controller;
    const { signal } = controller;
    const emptyText = t("commandPalette.empty.noMatchingCommands");
    const searchingText = t("commandPalette.searching");

    if (limits.maxResults <= 0) {
      renderGroups([], emptyText);
      chunkSearchController = null;
      return;
    }

    const compiled = compileFuzzyQuery(querySnapshot);
    const maxRanked = Math.min(cachedCommands.length, limits.maxResults * 3);
    const top: RenderableCommand[] = [];
    let cursor = 0;

    const isBetter = (a: RenderableCommand, b: RenderableCommand): boolean => {
      if (a.score !== b.score) return a.score > b.score;
      return a.title.localeCompare(b.title) < 0;
    };

    const worstIndex = (): number => {
      let worst = 0;
      for (let i = 1; i < top.length; i += 1) {
        if (isBetter(top[worst]!, top[i]!)) worst = i;
      }
      return worst;
    };

    const processChunk = (endExclusive: number): void => {
      for (let i = cursor; i < endExclusive; i += 1) {
        if (signal.aborted) return;
        const cmd = cachedCommands[i]!;
        const match = fuzzyMatchCommandPrepared(compiled, cmd);
        if (!match) continue;
        const candidate: RenderableCommand = { ...cmd, score: match.score, titleRanges: match.titleRanges };

        if (top.length < maxRanked) {
          top.push(candidate);
          continue;
        }

        const idx = worstIndex();
        if (isBetter(candidate, top[idx]!)) top[idx] = candidate;
      }
    };

    const makeGroupsFromTop = (): CommandGroup[] => {
      const ranked = [...top].sort((a, b) => {
        if (a.score !== b.score) return b.score - a.score;
        return a.title.localeCompare(b.title);
      });

      const groupsByLabel = new Map<string, RenderableCommand[]>();
      const order: string[] = [];
      let remainingSlots = limits.maxResults;
      for (const cmd of ranked) {
        if (remainingSlots <= 0) break;
        const label = groupLabel(cmd.category);
        if (!groupsByLabel.has(label)) order.push(label);
        const list = groupsByLabel.get(label) ?? [];
        if (list.length < limits.maxResultsPerGroup) {
          list.push(cmd);
          groupsByLabel.set(label, list);
          remainingSlots -= 1;
        }
      }
      return order.map((label) => ({ label, commands: groupsByLabel.get(label) ?? [] }));
    };

    const scheduleNext = () =>
      new Promise<void>((resolve) => {
        if (typeof requestAnimationFrame === "function") requestAnimationFrame(() => resolve());
        else setTimeout(resolve, 0);
      });

    const initialEnd = Math.min(cachedCommands.length, cursor + CHUNK_SEARCH_CHUNK_SIZE);
    processChunk(initialEnd);
    cursor = initialEnd;

    let hasRenderedAnyMatches = top.length > 0;
    if (!signal.aborted && isOpen && input.value === querySnapshot) {
      if (hasRenderedAnyMatches) {
        renderGroups(makeGroupsFromTop(), emptyText);
      } else {
        // Don't show a "no results" empty state until the scan completes. If the first
        // chunk doesn't contain matches, later chunks still might.
        renderGroups([], searchingText);
      }
    }

    void (async () => {
      try {
        while (!signal.aborted && cursor < cachedCommands.length) {
          // Yield to the browser so we don't monopolize the main thread for huge lists.
          await scheduleNext();
          if (signal.aborted) return;
          if (!isOpen) return;
          if (input.value !== querySnapshot) return;

          const end = Math.min(cachedCommands.length, cursor + CHUNK_SEARCH_CHUNK_SIZE);
          processChunk(end);
          cursor = end;

          if (!hasRenderedAnyMatches && top.length > 0) {
            hasRenderedAnyMatches = true;
            renderGroups(makeGroupsFromTop(), emptyText);
          }
        }

        if (signal.aborted) return;
        if (!isOpen) return;
        if (input.value !== querySnapshot) return;
        // Render final results.
        if (top.length === 0) {
          renderGroups([], emptyText);
        } else {
          renderGroups(makeGroupsFromTop(), emptyText);
        }
      } finally {
        if (chunkSearchController === controller) {
          chunkSearchController = null;
        }
      }
    })().catch(() => {});
  }

  function renderResults(mode: "sync" | "async"): void {
    ensureCommandsCache();

    const trimmed = query.trim();
    const shortcutMode = trimmed.startsWith("/");
    hint.hidden = !shortcutMode;

    if (shortcutMode) {
      abortChunkedSearch();
      const shortcutQuery = trimmed.slice(1).trim();
      const matches = searchShortcutCommands({
        commands: cachedCommands,
        keybindingIndex,
        query: shortcutQuery,
        limits: { maxResults: limits.maxResults, maxResultsPerCategory: limits.maxResultsPerGroup },
      });
      renderGroups(
        buildGroupsForShortcutMode(matches, limits),
        shortcutQuery ? t("commandPalette.empty.noMatchingShortcuts") : t("commandPalette.empty.noShortcuts"),
      );
      return;
    }

    if (!trimmed) {
      abortChunkedSearch();
      renderGroups(buildGroupsForEmptyQuery(cachedCommands, getRecentsForDisplay(cachedCommands), limits), t("commandPalette.empty.noCommands"));
      return;
    }

    const shouldChunk =
      mode === "async" && trimmed.length >= CHUNK_SEARCH_MIN_QUERY_LEN && cachedCommands.length >= CHUNK_SEARCH_MIN_COMMANDS;
    if (shouldChunk) {
      startChunkedSearch(query);
      return;
    }

    abortChunkedSearch();
    renderGroups(buildGroupsForQuery(cachedCommands, query, limits), t("commandPalette.empty.noMatchingCommands"));
  }

  const debouncedRender = debounce(
    () => {
      if (!isOpen) return;
      renderResults("async");
    },
    inputDebounce,
  );

  const prepareForKeyboardInteraction = (): void => {
    if (!isOpen) return;
    // If the user navigates/executes immediately after typing, ensure we apply the
    // latest query first. Prefer flushing the debounced render (which may start a
    // chunked search) rather than forcing a potentially expensive synchronous re-score.
    if (debouncedRender.pending()) {
      debouncedRender.flush();
    }

    // If we're still refining results in the background, freeze the list so keyboard
    // navigation doesn't jump around while the user is moving selection.
    if (chunkSearchController) {
      abortChunkedSearch();
    }
  };

  function updateSelection(nextIndex: number): void {
    if (visibleItems.length === 0) {
      selectedIndex = 0;
      return;
    }

    const prev = selectedIndex;
    selectedIndex = Math.max(0, Math.min(nextIndex, visibleItems.length - 1));
    if (prev === selectedIndex) return;

    const prevEl = visibleItemEls[prev];
    if (prevEl) prevEl.setAttribute("aria-selected", "false");

    const nextEl = visibleItemEls[selectedIndex];
    if (nextEl) {
      nextEl.setAttribute("aria-selected", "true");
      setActiveDescendant(nextEl.id);
      if (typeof nextEl.scrollIntoView === "function") {
        nextEl.scrollIntoView({ block: "nearest" });
      }
    } else {
      setActiveDescendant(null);
    }
  }

  function runSelected(): void {
    const item = visibleItems[selectedIndex];
    if (!item) return;
    close();
    if (item.kind === "goTo") {
      goTo?.onGoTo(item.parsed);
      return;
    }
    if (item.kind === "function") {
      onSelectFunction?.(item.name);
      return;
    }
    executeCommand(item.command.commandId);
  }

  const disposeRegistrySub = commandRegistry.subscribe(() => {
    commandsCacheDirty = true;
    if (!isOpen) return;
    debouncedRender.cancel();
    renderResults("async");
  });

  // Commands can be gated by context keys via `command.when` (e.g. permissions).
  // Keep the cached command list in sync with context changes so the palette
  // doesn't get stuck showing stale availability until the registry changes.
  const disposeContextSub = contextKeys.onDidChange(() => {
    commandsCacheDirty = true;
    if (!isOpen) return;
    // If we're currently refining results in the background, abort the stale search
    // and re-run against the updated command set.
    abortChunkedSearch();
    // Debounce so rapid context key changes (focus transitions, etc) don't cause
    // repeated full rescoring while the palette is open.
    debouncedRender();
  });

  const disposeRecentsTracker = installCommandPaletteRecentsTracking(commandRegistry, storage);

  const onOverlayClick = (e: MouseEvent) => {
    if (e.target === overlay) close();
  };
  overlay.addEventListener("click", onOverlayClick);

  const onInput = () => {
    query = input.value;
    selectedIndex = 0;
    abortChunkedSearch();
    debouncedRender();
  };
  input.addEventListener("input", onInput);

  const onInputKeyDown = (e: KeyboardEvent) => {
    if (e.key === "Escape") {
      e.preventDefault();
      close();
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      prepareForKeyboardInteraction();
      updateSelection(selectedIndex + 1);
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      prepareForKeyboardInteraction();
      updateSelection(selectedIndex - 1);
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      prepareForKeyboardInteraction();
      runSelected();
    }
  };
  input.addEventListener("keydown", onInputKeyDown);
  list.addEventListener("keydown", onInputKeyDown);

  function dispose(): void {
    // Treat disposal as a hard-close for context key purposes.
    contextKeys.set(COMMAND_PALETTE_OPEN_CONTEXT_KEY, false);
    isOpen = false;
    document.removeEventListener("focusin", handleDocumentFocusIn);
    debouncedRender.cancel();
    abortChunkedSearch();
    if (extensionLoadTimer != null) {
      window.clearTimeout(extensionLoadTimer);
      extensionLoadTimer = null;
    }
    disposeRegistrySub();
    disposeContextSub();
    disposeRecentsTracker();
    overlay.removeEventListener("click", onOverlayClick);
    overlay.removeEventListener("keydown", handleOverlayKeyDown);
    input.removeEventListener("input", onInput);
    input.removeEventListener("keydown", onInputKeyDown);
    list.removeEventListener("keydown", onInputKeyDown);
    overlay.remove();
  }

  return { open, close, isOpen: () => isOpen, dispose };
}
