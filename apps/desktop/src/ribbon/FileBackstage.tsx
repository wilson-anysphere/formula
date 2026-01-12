import React from "react";

import type { RibbonFileActions } from "./ribbonSchema.js";

export interface FileBackstageProps {
  open: boolean;
  actions?: RibbonFileActions;
  onClose: () => void;
}

type BackstageItem = {
  label: string;
  hint: string;
  testId: string;
  ariaLabel: string;
  onInvoke?: () => void;
};

export function FileBackstage({ open, actions, onClose }: FileBackstageProps) {
  const panelRef = React.useRef<HTMLDivElement | null>(null);
  const firstButtonRef = React.useRef<HTMLButtonElement | null>(null);

  const isMac = React.useMemo(() => {
    if (typeof navigator === "undefined") return false;
    return /Mac|iPhone|iPad|iPod/.test(navigator.platform);
  }, []);

  const shortcut = React.useCallback(
    (key: string, options: { shift?: boolean } = {}) => {
      if (isMac) {
        return `${options.shift ? "⌘⇧" : "⌘"}${key}`;
      }
      return `Ctrl+${options.shift ? "Shift+" : ""}${key}`;
    },
    [isMac],
  );

  const items = React.useMemo<BackstageItem[]>(
    () => [
      {
        label: "New Workbook",
        hint: shortcut("N"),
        testId: "file-new",
        ariaLabel: "New workbook",
        onInvoke: actions?.newWorkbook,
      },
      { label: "Open…", hint: shortcut("O"), testId: "file-open", ariaLabel: "Open workbook", onInvoke: actions?.openWorkbook },
      { label: "Save", hint: shortcut("S"), testId: "file-save", ariaLabel: "Save workbook", onInvoke: actions?.saveWorkbook },
      {
        label: "Save As…",
        hint: shortcut("S", { shift: true }),
        testId: "file-save-as",
        ariaLabel: "Save workbook as",
        onInvoke: actions?.saveWorkbookAs,
      },
      { label: "Close Window", hint: shortcut("W"), testId: "file-close", ariaLabel: "Close window", onInvoke: actions?.closeWindow },
      { label: "Quit", hint: shortcut("Q"), testId: "file-quit", ariaLabel: "Quit application", onInvoke: actions?.quit },
    ],
    [actions, shortcut],
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

  const trapTab = React.useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== "Tab") return;
    const panel = panelRef.current;
    if (!panel) return;
    const focusables = Array.from(panel.querySelectorAll<HTMLElement>("button:not([disabled]), [href], input, select, textarea, [tabindex]")).filter(
      (el) => el.getAttribute("aria-hidden") !== "true" && !el.hasAttribute("disabled"),
    );
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
                disabled={disabled}
                onClick={() => {
                  onClose();
                  item.onInvoke?.();
                }}
              >
                <span className="ribbon-backstage__label">{item.label}</span>
                <span className="ribbon-backstage__hint">{item.hint}</span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}
