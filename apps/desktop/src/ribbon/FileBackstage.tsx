import React from "react";

import type { RibbonFileActions } from "./ribbonSchema.js";
import { RibbonIcon, type RibbonIconId } from "./icons/index.js";
import { getRibbonUiStateSnapshot, subscribeRibbonUiState } from "./ribbonUiState.js";

export interface FileBackstageProps {
  open: boolean;
  actions?: RibbonFileActions;
  onClose: () => void;
}

type BackstageItem = {
  iconId: RibbonIconId;
  label: string;
  hint: string;
  ariaKeyShortcuts: string;
  testId: string;
  ariaLabel: string;
  onInvoke?: () => void;
  kind?: "command" | "toggle";
  pressed?: boolean;
};

export function FileBackstage({ open, actions, onClose }: FileBackstageProps) {
  const panelRef = React.useRef<HTMLDivElement | null>(null);
  const firstButtonRef = React.useRef<HTMLButtonElement | null>(null);

  const uiState = React.useSyncExternalStore(subscribeRibbonUiState, getRibbonUiStateSnapshot, getRibbonUiStateSnapshot);
  const autoSavePressed = Boolean(uiState.pressedById["file.save.autoSave"]);

  const isMac = React.useMemo(() => {
    if (typeof navigator === "undefined") return false;
    return /Mac|iPhone|iPad|iPod/.test(navigator.platform);
  }, []);

  const shortcut = React.useCallback(
    (commandId: string, fallbackKey: string, options: { shift?: boolean } = {}) => {
      const fromIndex = uiState.shortcutById?.[commandId];
      if (fromIndex) return fromIndex;
      if (isMac) return `${options.shift ? "⇧" : ""}⌘${fallbackKey}`;
      return `Ctrl+${options.shift ? "Shift+" : ""}${fallbackKey}`;
    },
    [isMac, uiState.shortcutById],
  );

  const ariaShortcut = React.useCallback(
    (commandId: string, fallbackKey: string, options: { shift?: boolean } = {}) => {
      const fromIndex = uiState.ariaKeyShortcutsById?.[commandId];
      if (fromIndex) return fromIndex;
      const parts: string[] = [];
      if (isMac) {
        if (options.shift) parts.push("Shift");
        parts.push("Meta");
      } else {
        parts.push("Control");
        if (options.shift) parts.push("Shift");
      }
      parts.push(fallbackKey.toUpperCase());
      return parts.join("+");
    },
    [isMac, uiState.ariaKeyShortcutsById],
  );

  const items = React.useMemo<BackstageItem[]>(
    () => [
      {
        iconId: "filePlus",
        label: "New Workbook",
        hint: shortcut("workbench.newWorkbook", "N"),
        ariaKeyShortcuts: ariaShortcut("workbench.newWorkbook", "N"),
        testId: "file-new",
        ariaLabel: "New workbook",
        onInvoke: actions?.newWorkbook,
      },
      {
        iconId: "folderOpen",
        label: "Open…",
        hint: shortcut("workbench.openWorkbook", "O"),
        ariaKeyShortcuts: ariaShortcut("workbench.openWorkbook", "O"),
        testId: "file-open",
        ariaLabel: "Open workbook",
        onInvoke: actions?.openWorkbook,
      },
      {
        iconId: "save",
        label: "Save",
        hint: shortcut("workbench.saveWorkbook", "S"),
        ariaKeyShortcuts: ariaShortcut("workbench.saveWorkbook", "S"),
        testId: "file-save",
        ariaLabel: "Save workbook",
        onInvoke: actions?.saveWorkbook,
      },
      {
        iconId: "edit",
        label: "Save As…",
        hint: shortcut("workbench.saveWorkbookAs", "S", { shift: true }),
        ariaKeyShortcuts: ariaShortcut("workbench.saveWorkbookAs", "S", { shift: true }),
        testId: "file-save-as",
        ariaLabel: "Save workbook as",
        onInvoke: actions?.saveWorkbookAs,
      },
      {
        iconId: "cloud",
        label: "AutoSave",
        hint: autoSavePressed ? "On" : "Off",
        ariaKeyShortcuts: "",
        testId: "file-auto-save",
        ariaLabel: "Toggle AutoSave",
        kind: "toggle",
        pressed: autoSavePressed,
        onInvoke: actions?.toggleAutoSave ? () => actions.toggleAutoSave?.(!autoSavePressed) : undefined,
      },
      {
        iconId: "clock",
        label: "Version History",
        hint: "",
        ariaKeyShortcuts: "",
        testId: "file-version-history",
        ariaLabel: "Version history",
        onInvoke: actions?.versionHistory,
      },
      {
        iconId: "shuffle",
        label: "Branches",
        hint: "",
        ariaKeyShortcuts: "",
        testId: "file-branch-manager",
        ariaLabel: "Branch manager",
        onInvoke: actions?.branchManager,
      },
      {
        iconId: "print",
        label: "Print…",
        hint: shortcut("workbench.print", "P"),
        ariaKeyShortcuts: ariaShortcut("workbench.print", "P"),
        testId: "file-print",
        ariaLabel: "Print",
        onInvoke: actions?.print,
      },
      {
        iconId: "eye",
        label: "Print Preview",
        hint: "",
        ariaKeyShortcuts: "",
        testId: "file-print-preview",
        ariaLabel: "Print preview",
        onInvoke: actions?.printPreview,
      },
      {
        iconId: "settings",
        label: "Page Setup…",
        hint: "",
        ariaKeyShortcuts: "",
        testId: "file-page-setup",
        ariaLabel: "Page setup",
        onInvoke: actions?.pageSetup,
      },
      {
        iconId: "close",
        label: "Close Window",
        hint: shortcut("workbench.closeWorkbook", "W"),
        ariaKeyShortcuts: ariaShortcut("workbench.closeWorkbook", "W"),
        testId: "file-close",
        ariaLabel: "Close window",
        onInvoke: actions?.closeWindow,
      },
      {
        iconId: "close",
        label: "Quit",
        hint: shortcut("workbench.quit", "Q"),
        ariaKeyShortcuts: ariaShortcut("workbench.quit", "Q"),
        testId: "file-quit",
        ariaLabel: "Quit application",
        onInvoke: actions?.quit,
      },
    ],
    [actions, ariaShortcut, autoSavePressed, shortcut],
  );

  const focusFirst = React.useCallback(() => {
    const focusables = panelRef.current?.querySelectorAll<HTMLButtonElement>("button:not([disabled])") ?? [];
    const fallback = firstButtonRef.current;
    const first = focusables[0] ?? fallback;
    first?.focus();
  }, []);

  React.useEffect(() => {
    if (!open) return;
    // Defer so the overlay is painted before we move focus.
    requestAnimationFrame(() => focusFirst());
  }, [focusFirst, open]);

  const moveFocus = React.useCallback((direction: "next" | "prev") => {
    const panel = panelRef.current;
    if (!panel) return;
    const focusables = Array.from(panel.querySelectorAll<HTMLButtonElement>("button:not([disabled])"));
    if (focusables.length === 0) return;
    const active = document.activeElement as HTMLElement | null;
    const currentIndex = active ? focusables.findIndex((el) => el === active) : -1;
    const delta = direction === "next" ? 1 : -1;
    const nextIndex = currentIndex >= 0 ? (currentIndex + delta + focusables.length) % focusables.length : 0;
    focusables[nextIndex]?.focus();
  }, []);

  const trapTab = React.useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== "Tab") return;
    const panel = panelRef.current;
    if (!panel) return;
    const focusables = Array.from(
      panel.querySelectorAll<HTMLElement>(
        'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
      ),
    ).filter((el) => el.getAttribute("aria-hidden") !== "true");
    if (focusables.length === 0) return;
    const first = focusables[0]!;
    const last = focusables[focusables.length - 1]!;
    const active = document.activeElement as HTMLElement | null;
    if (!active) return;

    if (event.shiftKey) {
      if (active === first) {
        event.preventDefault();
        last.focus();
      }
      return;
    }

    if (active === last) {
      event.preventDefault();
      first.focus();
    }
  }, []);

  if (!open) return null;

  return (
    <div
      className="ribbon-backstage-overlay"
      data-keybinding-barrier="true"
      role="dialog"
      aria-modal="true"
      aria-label="File menu"
      onMouseDown={(event) => {
        if (event.target !== event.currentTarget) return;
        onClose();
      }}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          onClose();
          return;
        }
        if (event.key === "ArrowDown") {
          event.preventDefault();
          moveFocus("next");
          return;
        }
        if (event.key === "ArrowUp") {
          event.preventDefault();
          moveFocus("prev");
          return;
        }
        trapTab(event);
      }}
    >
      <div ref={panelRef} className="ribbon-backstage">
        <div className="ribbon-backstage__title">File</div>
        <div className="ribbon-backstage__list" role="menu" aria-label="File actions">
          {items.map((item, idx) => {
            const disabled = !item.onInvoke;
            return (
              <button
                // eslint-disable-next-line react/no-array-index-key
                key={`${item.testId}-${idx}`}
                ref={idx === 0 ? firstButtonRef : undefined}
                type="button"
                className="ribbon-backstage__item"
                data-testid={item.testId}
                aria-label={item.ariaLabel}
                aria-keyshortcuts={item.ariaKeyShortcuts || undefined}
                role={item.kind === "toggle" ? "menuitemcheckbox" : "menuitem"}
                aria-checked={item.kind === "toggle" ? Boolean(item.pressed) : undefined}
                aria-pressed={item.kind === "toggle" ? Boolean(item.pressed) : undefined}
                disabled={disabled}
                onClick={() => {
                  onClose();
                  item.onInvoke?.();
                }}
              >
                <span className="ribbon-backstage__item-main">
                  <span className="ribbon-backstage__icon" aria-hidden="true">
                    <RibbonIcon id={item.iconId} />
                  </span>
                  <span className="ribbon-backstage__label">{item.label}</span>
                </span>
                <span className="ribbon-backstage__hint">{item.hint}</span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
