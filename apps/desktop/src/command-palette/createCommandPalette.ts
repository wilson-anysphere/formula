import type { CommandContribution, CommandRegistry } from "../extensions/commandRegistry.js";
import type { ContextKeyService } from "../extensions/contextKeys.js";

import { fuzzyMatchCommand, type MatchRange } from "./fuzzy.js";
import { readCommandPaletteRecents, recordCommandPaletteRecent } from "./recents.js";

type RenderableCommand = CommandContribution & {
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
  keybindingIndex: Map<string, readonly string[]>;
  ensureExtensionsLoaded: () => Promise<void>;
  onCloseFocus: () => void;
  placeholder?: string;
  /**
   * How long to wait after opening before we kick off extension loading.
   * Defaults to 600ms to avoid paying the extension-worker cost for "quick" palette usage
   * (notably in e2e tests that just invoke built-in commands).
   */
  extensionLoadDelayMs?: number;
};

export type CommandPaletteController = {
  open: () => void;
  close: () => void;
  dispose: () => void;
};

function groupLabel(category: string | null): string {
  const value = String(category ?? "").trim();
  return value ? value : "Other";
}

function sortCommandsAlpha(a: CommandContribution, b: CommandContribution): number {
  return a.title.localeCompare(b.title);
}

function buildGroupsForEmptyQuery(allCommands: CommandContribution[], recentIds: string[]): CommandGroup[] {
  const byId = new Map(allCommands.map((cmd) => [cmd.commandId, cmd]));

  const recents: RenderableCommand[] = [];
  const recentSet = new Set<string>();
  for (const id of recentIds) {
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
  if (recents.length > 0) groups.push({ label: "RECENT", commands: recents });

  const sortedCategoryLabels = [...categories.keys()].sort((a, b) => a.localeCompare(b));
  for (const label of sortedCategoryLabels) {
    const cmds = categories.get(label)!;
    cmds.sort(sortCommandsAlpha);
    groups.push({ label, commands: cmds });
  }

  return groups;
}

function buildGroupsForQuery(allCommands: CommandContribution[], query: string): CommandGroup[] {
  const matches: RenderableCommand[] = [];
  for (const cmd of allCommands) {
    const match = fuzzyMatchCommand(query, cmd);
    if (!match) continue;
    matches.push({ ...cmd, score: match.score, titleRanges: match.titleRanges });
  }

  matches.sort((a, b) => {
    if (a.score !== b.score) return b.score - a.score;
    return a.title.localeCompare(b.title);
  });

  const groupsByLabel = new Map<string, RenderableCommand[]>();
  const groupOrder: string[] = [];

  for (const cmd of matches) {
    const label = groupLabel(cmd.category);
    if (!groupsByLabel.has(label)) groupOrder.push(label);
    const list = groupsByLabel.get(label) ?? [];
    list.push(cmd);
    groupsByLabel.set(label, list);
  }

  return groupOrder.map((label) => ({ label, commands: groupsByLabel.get(label) ?? [] }));
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
    placeholder = "Type a command…",
    extensionLoadDelayMs = 600,
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

  const palette = document.createElement("div");
  palette.className = "command-palette";
  palette.dataset.testid = "command-palette";

  const input = document.createElement("input");
  input.className = "command-palette__input";
  input.dataset.testid = "command-palette-input";
  input.placeholder = placeholder;

  const list = document.createElement("ul");
  list.className = "command-palette__list";
  list.dataset.testid = "command-palette-list";

  palette.appendChild(input);
  palette.appendChild(list);
  overlay.appendChild(palette);
  document.body.appendChild(overlay);

  let isOpen = false;
  let query = "";
  let selectedIndex = 0;
  let visibleCommands: RenderableCommand[] = [];
  let extensionLoadTimer: number | null = null;

  const executeCommand = (commandId: string): void => {
    void commandRegistry.executeCommand(commandId).catch((err) => {
      console.error(`Command failed (${commandId}):`, err);
    });
  };

  function close(): void {
    if (!isOpen) return;
    isOpen = false;
    overlay.style.display = "none";
    query = "";
    selectedIndex = 0;
    visibleCommands = [];
    if (extensionLoadTimer != null) {
      window.clearTimeout(extensionLoadTimer);
      extensionLoadTimer = null;
    }
    onCloseFocus();
  }

  function open(): void {
    query = "";
    selectedIndex = 0;
    input.value = "";
    overlay.style.display = "flex";
    isOpen = true;

    render();

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

  function getRecentsForDisplay(allCommands: CommandContribution[]): string[] {
    const ids = readCommandPaletteRecents(localStorage);
    if (ids.length === 0) return [];
    const existing = new Set(allCommands.map((cmd) => cmd.commandId));
    return ids.filter((id) => existing.has(id));
  }

  function render(): void {
    const allCommands = commandRegistry.listCommands();

    const groups =
      query.trim() === ""
        ? buildGroupsForEmptyQuery(allCommands, getRecentsForDisplay(allCommands))
        : buildGroupsForQuery(allCommands, query);

    visibleCommands = groups.flatMap((g) => g.commands);
    if (selectedIndex >= visibleCommands.length) selectedIndex = Math.max(0, visibleCommands.length - 1);

    list.replaceChildren();

    if (groups.length === 0 || visibleCommands.length === 0) {
      const empty = document.createElement("li");
      empty.className = "command-palette__empty";
      empty.textContent = query.trim() ? "No matching commands" : "No commands";
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
        main.appendChild(renderHighlightedText(cmd.title, cmd.titleRanges));

        const shortcut = keybindingIndex.get(cmd.commandId)?.[0] ?? null;
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
          recordCommandPaletteRecent(localStorage, cmd.commandId);
          executeCommand(cmd.commandId);
        });

        list.appendChild(li);

        if (globalIndex === selectedIndex) {
          // Keep selection in view when navigating via keyboard.
          queueMicrotask(() => li.scrollIntoView({ block: "nearest" }));
        }
      }

      commandOffset += group.commands.length;
    }
  }

  function runSelected(): void {
    const cmd = visibleCommands[selectedIndex];
    if (!cmd) return;
    close();
    recordCommandPaletteRecent(localStorage, cmd.commandId);
    executeCommand(cmd.commandId);
  }

  const disposeRegistrySub = commandRegistry.subscribe(() => {
    if (!isOpen) return;
    render();
  });

  const onOverlayClick = (e: MouseEvent) => {
    if (e.target === overlay) close();
  };
  overlay.addEventListener("click", onOverlayClick);

  const onInput = () => {
    query = input.value;
    selectedIndex = 0;
    render();
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
      selectedIndex = visibleCommands.length === 0 ? 0 : Math.min(visibleCommands.length - 1, selectedIndex + 1);
      render();
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      selectedIndex = visibleCommands.length === 0 ? 0 : Math.max(0, selectedIndex - 1);
      render();
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
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
    disposeRegistrySub();
    overlay.removeEventListener("click", onOverlayClick);
    input.removeEventListener("input", onInput);
    input.removeEventListener("keydown", onInputKeyDown);
    window.removeEventListener("keydown", onGlobalKeyDown);
    overlay.remove();
  }

  return { open, close, dispose };
}
