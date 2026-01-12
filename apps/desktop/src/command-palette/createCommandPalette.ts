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
  let chunkSearchController: AbortController | null = null;
  let extensionLoadTimer: number | null = null;

  const executeCommand = (commandId: string): void => {
    void commandRegistry.executeCommand(commandId).catch((err) => {
      console.error(`Command failed (${commandId}):`, err);
    });
  };

  const abortChunkedSearch = (): void => {
    chunkSearchController?.abort();
    chunkSearchController = null;
  };

  function close(): void {
    if (!isOpen) return;
    isOpen = false;
    debouncedRender.cancel();
    abortChunkedSearch();
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
    abortChunkedSearch();
    query = "";
    selectedIndex = 0;
    input.value = "";
    overlay.style.display = "flex";
    isOpen = true;

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

  const commandRowCache = new Map<
    string,
    {
      li: HTMLLIElement;
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
    visibleCommands = groups.flatMap((g) => g.commands);
    if (selectedIndex >= visibleCommands.length) selectedIndex = Math.max(0, visibleCommands.length - 1);

    list.replaceChildren();
    visibleCommandEls = [];

    if (groups.length === 0 || visibleCommands.length === 0) {
      const empty = document.createElement("li");
      empty.className = "command-palette__empty";
      empty.textContent = emptyText;
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

        let cached = commandRowCache.get(cmd.commandId);
        if (!cached) {
          const li = document.createElement("li");
          li.className = "command-palette__item";

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

          cached = { li, label, description, right, shortcutPill };
          commandRowCache.set(cmd.commandId, cached);
        }

        cached.li.setAttribute("aria-selected", globalIndex === selectedIndex ? "true" : "false");
        cached.label.replaceChildren(renderHighlightedText(cmd.title, cmd.titleRanges));

        const descriptionText = typeof cmd.description === "string" ? cmd.description.trim() : "";
        cached.description.textContent = descriptionText;
        cached.description.style.display = descriptionText ? "" : "none";

        const kbValue = keybindingIndex.get(cmd.commandId);
        const shortcut =
          typeof kbValue === "string" ? kbValue : Array.isArray(kbValue) ? (kbValue[0] ?? null) : null;
        if (shortcut) {
          cached.shortcutPill.textContent = shortcut;
          cached.right.style.display = "";
        } else {
          cached.shortcutPill.textContent = "";
          cached.right.style.display = "none";
        }

        list.appendChild(cached.li);
        visibleCommandEls[globalIndex] = cached.li;
      }

      commandOffset += group.commands.length;
    }

    // Keep selection in view after re-render.
    queueMicrotask(() => visibleCommandEls[selectedIndex]?.scrollIntoView({ block: "nearest" }));
  }

  function startChunkedSearch(querySnapshot: string): void {
    ensureCommandsCache();

    abortChunkedSearch();
    const controller = new AbortController();
    chunkSearchController = controller;
    const { signal } = controller;
    const emptyText = t("commandPalette.empty.noMatchingCommands");

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

    if (!signal.aborted && isOpen && input.value === querySnapshot) {
      renderGroups(makeGroupsFromTop(), emptyText);
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
        }

        if (signal.aborted) return;
        if (!isOpen) return;
        if (input.value !== querySnapshot) return;
        // Render final results.
        renderGroups(makeGroupsFromTop(), emptyText);
      } finally {
        if (chunkSearchController === controller) {
          chunkSearchController = null;
        }
      }
    })();
  }

  function renderResults(mode: "sync" | "async"): void {
    ensureCommandsCache();

    const trimmed = query.trim();
    const shortcutMode = trimmed.startsWith("/");
    hint.hidden = !shortcutMode;

    if (shortcutMode) {
      abortChunkedSearch();
      const shortcutQuery = trimmed.slice(1).trim();
      const matches = searchShortcutCommands({ commands: cachedCommands, keybindingIndex, query: shortcutQuery });
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
    renderResults("async");
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
    abortChunkedSearch();
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
