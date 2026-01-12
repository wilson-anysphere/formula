import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import type { ContextKeyService } from "../extensions/contextKeys.js";

import { t } from "../i18n/index.js";

import { debounce } from "./debounce.js";
import {
  compileFuzzyQuery,
  fuzzyMatchCommandPrepared,
  prepareCommandForFuzzy,
  type MatchRange,
  type PreparedCommandForFuzzy,
} from "./fuzzy.js";
import { getRecentCommandIdsForDisplay, installCommandRecentsTracker } from "./recents.js";
import { searchShortcutCommands } from "./shortcutSearch.js";

type RenderableCommand = PreparedCommandForFuzzy<CommandContribution> & {
  score: number;
  titleRanges: MatchRange[];
};

type CommandGroup = {
  label: string;
  commands: RenderableCommand[];
};

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
};

export type CommandPaletteController = {
  open: () => void;
  close: () => void;
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

  const categories = new Map<string, RenderableCommand[]>();
  for (const cmd of remaining) {
    const label = groupLabel(cmd.category);
    const list = categories.get(label) ?? [];
    list.push({ ...cmd, score: 0, titleRanges: [] });
    categories.set(label, list);
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
  const compiled = compileFuzzyQuery(query);

  // Keep only top N matches to avoid sorting huge arrays (N = maxResults).
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
    if (top.length < limits.maxResults) {
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

  for (const cmd of top) {
    const label = groupLabel(cmd.category);
    if (!groupsByLabel.has(label)) groupOrder.push(label);
    const list = groupsByLabel.get(label) ?? [];
    if (list.length < limits.maxResultsPerGroup) {
      list.push(cmd);
      groupsByLabel.set(label, list);
    }
  }

  return groupOrder.map((label) => ({ label, commands: groupsByLabel.get(label) ?? [] }));
}

function buildGroupsForShortcutMode(matches: Array<PreparedCommandForFuzzy<CommandContribution> & { shortcut: string }>): CommandGroup[] {
  const groupsByLabel = new Map<string, RenderableCommand[]>();
  const order: string[] = [];

  for (const cmd of matches) {
    const label = groupLabel(cmd.category);
    if (!groupsByLabel.has(label)) order.push(label);
    const list = groupsByLabel.get(label) ?? [];
    list.push({ ...cmd, score: 0, titleRanges: [] });
    groupsByLabel.set(label, list);
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

export function createCommandPalette(options: CreateCommandPaletteOptions): CommandPaletteController {
  const {
    commandRegistry,
    contextKeys: _contextKeys,
    keybindingIndex,
    ensureExtensionsLoaded,
    onCloseFocus,
    placeholder = t("commandPalette.placeholder"),
    extensionLoadDelayMs = 600,
    maxResults = 100,
    maxResultsPerGroup = 20,
    inputDebounceMs = 70,
  } = options;

  const overlay = document.createElement("div");
  overlay.style.position = "fixed";
  overlay.style.inset = "0";
  overlay.style.display = "none";
  overlay.style.alignItems = "flex-start";
  overlay.style.justifyContent = "center";
  overlay.style.paddingTop = "80px";
  overlay.style.background = "var(--dialog-backdrop)";
  overlay.style.zIndex = "1000";
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");

  const palette = document.createElement("div");
  palette.className = "command-palette";
  palette.dataset.testid = "command-palette";

  const input = document.createElement("input");
  input.className = "command-palette__input";
  input.dataset.testid = "command-palette-input";
  input.placeholder = placeholder;

  const hint = document.createElement("div");
  hint.className = "command-palette__hint";
  hint.textContent = t("commandPalette.shortcutSearch.hint");
  hint.hidden = true;

  const list = document.createElement("ul");
  list.className = "command-palette__list";
  list.dataset.testid = "command-palette-list";

  palette.appendChild(input);
  palette.appendChild(hint);
  palette.appendChild(list);
  overlay.appendChild(palette);
  document.body.appendChild(overlay);

  const limits = {
    maxResults: Math.max(0, Math.floor(maxResults)),
    maxResultsPerGroup: Math.max(1, Math.floor(maxResultsPerGroup)),
  };
  const inputDebounce = Math.max(0, Math.floor(inputDebounceMs));

  let isOpen = false;
  let query = "";
  let selectedIndex = 0;
  let visibleCommands: RenderableCommand[] = [];
  let visibleCommandEls: HTMLLIElement[] = [];
  let commandsCacheDirty = true;
  let cachedCommands: PreparedCommandForFuzzy<CommandContribution>[] = [];
  let extensionLoadTimer: number | null = null;

  const executeCommand = (commandId: string): void => {
    void commandRegistry.executeCommand(commandId).catch((err) => {
      console.error(`Command failed (${commandId}):`, err);
    });
  };

  function close(): void {
    if (!isOpen) return;
    isOpen = false;
    debouncedRender.cancel();
    overlay.style.display = "none";
    query = "";
    selectedIndex = 0;
    visibleCommands = [];
    visibleCommandEls = [];
    if (extensionLoadTimer != null) {
      window.clearTimeout(extensionLoadTimer);
      extensionLoadTimer = null;
    }
    onCloseFocus();
  }

  function open(): void {
    debouncedRender.cancel();
    query = "";
    selectedIndex = 0;
    input.value = "";
    overlay.style.display = "flex";
    isOpen = true;

    renderResults();

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
    cachedCommands = commandRegistry
      .listCommands()
      // The command palette owns opening itself; avoid showing a no-op entry.
      .filter((cmd) => cmd.commandId !== "workbench.showCommandPalette")
      .map((cmd) => prepareCommandForFuzzy(cmd));
    commandsCacheDirty = false;
  }

  function getRecentsForDisplay(allCommands: Array<{ commandId: string }>): string[] {
    return getRecentCommandIdsForDisplay(
      localStorage,
      allCommands.map((cmd) => cmd.commandId),
    );
  }

  function renderResults(): void {
    ensureCommandsCache();

    const trimmed = query.trim();
    const shortcutMode = trimmed.startsWith("/");
    hint.hidden = !shortcutMode;

    const shortcutQuery = shortcutMode ? trimmed.slice(1).trim() : "";
    const groups = shortcutMode
      ? buildGroupsForShortcutMode(
          searchShortcutCommands({
            commands: cachedCommands,
            keybindingIndex,
            query: shortcutQuery,
          }),
        )
      : trimmed === ""
        ? buildGroupsForEmptyQuery(cachedCommands, getRecentsForDisplay(cachedCommands), limits)
        : buildGroupsForQuery(cachedCommands, query, limits);

    visibleCommands = groups.flatMap((g) => g.commands);
    if (selectedIndex >= visibleCommands.length) selectedIndex = Math.max(0, visibleCommands.length - 1);

    list.replaceChildren();
    visibleCommandEls = [];

    if (groups.length === 0 || visibleCommands.length === 0) {
      const empty = document.createElement("li");
      empty.className = "command-palette__empty";
      if (shortcutMode) {
        empty.textContent = shortcutQuery
          ? t("commandPalette.empty.noMatchingShortcuts")
          : t("commandPalette.empty.noShortcuts");
      } else {
        empty.textContent = trimmed
          ? t("commandPalette.empty.noMatchingCommands")
          : t("commandPalette.empty.noCommands");
      }
      list.appendChild(empty);
      return;
    }

    let commandOffset = 0;
    for (const group of groups) {
      const header = document.createElement("li");
      header.className = "command-palette__group";
      header.textContent = group.label;
      list.appendChild(header);

      for (let i = 0; i < group.commands.length; i += 1) {
        const cmd = group.commands[i]!;
        const globalIndex = commandOffset + i;

         const li = document.createElement("li");
         li.className = "command-palette__item";
         li.setAttribute("aria-selected", globalIndex === selectedIndex ? "true" : "false");

         const main = document.createElement("div");
         main.className = "command-palette__item-main";

         const label = document.createElement("div");
         label.className = "command-palette__item-label";
         label.appendChild(renderHighlightedText(cmd.title, cmd.titleRanges));
         main.appendChild(label);

         const descriptionText = typeof cmd.description === "string" ? cmd.description.trim() : "";
         if (descriptionText) {
           const description = document.createElement("div");
           description.className = "command-palette__item-description";
           description.textContent = descriptionText;
           main.appendChild(description);
         }

        const kbValue = keybindingIndex.get(cmd.commandId);
        const shortcut =
          typeof kbValue === "string" ? kbValue : Array.isArray(kbValue) ? (kbValue[0] ?? null) : null;
        if (shortcut) {
          const right = document.createElement("div");
          right.className = "command-palette__item-right";
          const pill = document.createElement("span");
          pill.className = "command-palette__shortcut";
          pill.textContent = shortcut;
          right.appendChild(pill);
          li.appendChild(main);
          li.appendChild(right);
        } else {
          li.appendChild(main);
        }

        li.addEventListener("mousedown", (e) => {
          // Prevent focus leaving the input before we run the command.
          e.preventDefault();
        });
        li.addEventListener("click", () => {
          close();
          executeCommand(cmd.commandId);
        });

        list.appendChild(li);
        visibleCommandEls[globalIndex] = li;
      }

      commandOffset += group.commands.length;
    }

    // Keep selection in view after re-render.
    queueMicrotask(() => visibleCommandEls[selectedIndex]?.scrollIntoView({ block: "nearest" }));
  }

  const debouncedRender = debounce(
    () => {
      if (!isOpen) return;
      renderResults();
    },
    inputDebounce,
  );

  function updateSelection(nextIndex: number): void {
    if (visibleCommands.length === 0) {
      selectedIndex = 0;
      return;
    }

    const prev = selectedIndex;
    selectedIndex = Math.max(0, Math.min(nextIndex, visibleCommands.length - 1));
    if (prev === selectedIndex) return;

    const prevEl = visibleCommandEls[prev];
    if (prevEl) prevEl.setAttribute("aria-selected", "false");

    const nextEl = visibleCommandEls[selectedIndex];
    if (nextEl) {
      nextEl.setAttribute("aria-selected", "true");
      nextEl.scrollIntoView({ block: "nearest" });
    }
  }

  function runSelected(): void {
    const cmd = visibleCommands[selectedIndex];
    if (!cmd) return;
    close();
    executeCommand(cmd.commandId);
  }

  const disposeRegistrySub = commandRegistry.subscribe(() => {
    commandsCacheDirty = true;
    if (!isOpen) return;
    debouncedRender.cancel();
    renderResults();
  });

  const disposeRecentsTracker = installCommandRecentsTracker(commandRegistry, localStorage, {
    ignoreCommandIds: ["workbench.showCommandPalette"],
  });

  const onOverlayClick = (e: MouseEvent) => {
    if (e.target === overlay) close();
  };
  overlay.addEventListener("click", onOverlayClick);

  const onInput = () => {
    query = input.value;
    selectedIndex = 0;
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
      debouncedRender.flush();
      updateSelection(selectedIndex + 1);
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      debouncedRender.flush();
      updateSelection(selectedIndex - 1);
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      debouncedRender.flush();
      runSelected();
    }
  };
  input.addEventListener("keydown", onInputKeyDown);

  const onGlobalKeyDown = (e: KeyboardEvent) => {
    if (e.defaultPrevented) return;
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || !e.shiftKey) return;
    if (e.key !== "P" && e.key !== "p") return;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return;
    }

    e.preventDefault();
    open();
  };
  window.addEventListener("keydown", onGlobalKeyDown);

  function dispose(): void {
    debouncedRender.cancel();
    if (extensionLoadTimer != null) {
      window.clearTimeout(extensionLoadTimer);
      extensionLoadTimer = null;
    }
    disposeRegistrySub();
    disposeRecentsTracker();
    overlay.removeEventListener("click", onOverlayClick);
    input.removeEventListener("input", onInput);
    input.removeEventListener("keydown", onInputKeyDown);
    window.removeEventListener("keydown", onGlobalKeyDown);
    overlay.remove();
  }

  return { open, close, dispose };
}
