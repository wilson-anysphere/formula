import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { ContextMenu, type ContextMenuItem } from "../menus/contextMenu.js";
import { normalizeExcelColorToCss } from "../shared/colors.js";
import { resolveCssVar } from "../theme/cssVars.js";
import * as nativeDialogs from "../tauri/nativeDialogs";
import { validateSheetName, type SheetMeta, type TabColor, type WorkbookSheetStore } from "./workbookSheetStore";
import { computeWorkbookSheetMoveIndex } from "./sheetReorder";
import { showQuickPick } from "../extensions/ui.js";

const SHEET_TAB_COLOR_PALETTE: Array<{ label: string; token: string; fallbackCss: string }> = [
  { label: "Red", token: "--sheet-tab-red", fallbackCss: "red" },
  { label: "Orange", token: "--sheet-tab-orange", fallbackCss: "orange" },
  { label: "Yellow", token: "--sheet-tab-yellow", fallbackCss: "yellow" },
  { label: "Green", token: "--sheet-tab-green", fallbackCss: "green" },
  { label: "Blue", token: "--sheet-tab-blue", fallbackCss: "blue" },
  { label: "Purple", token: "--sheet-tab-purple", fallbackCss: "purple" },
  { label: "Gray", token: "--sheet-tab-gray", fallbackCss: "gray" },
];

type Props = {
  store: WorkbookSheetStore;
  activeSheetId: string;
  onActivateSheet: (sheetId: string) => void;
  onAddSheet: () => Promise<void> | void;
  /**
   * Optional hook to persist sheet renames before updating the local sheet metadata store.
   *
   * Used by the desktop (Tauri) shell to route renames through the backend so the workbook
   * stays consistent across reloads.
   */
  onPersistSheetRename?: (sheetId: string, name: string) => Promise<void> | void;
  /**
   * Optional hook to persist sheet deletion before updating the local sheet metadata store.
   *
   * Used by the desktop (Tauri) shell to route deletes through the backend so the workbook
   * stays consistent across reloads.
   */
  onPersistSheetDelete?: (sheetId: string) => Promise<void> | void;
  /**
   * Called after a sheet tab reorder is committed (drag-and-drop).
   *
   * Used by `main.ts` to restore focus back to the grid so users can continue
   * editing after reordering.
   */
  onSheetsReordered?: () => void;
  /**
   * Called after a sheet rename is successfully committed.
   *
   * Used by `main.ts` to rewrite DocumentController formulas referencing the old sheet name.
   */
  onSheetRenamed?: (event: { sheetId: string; oldName: string; newName: string }) => void;
  /**
   * Called after a sheet is successfully deleted from the metadata store.
   *
   * Used by `main.ts` to rewrite DocumentController formulas referencing the deleted sheet name.
   */
  onSheetDeleted?: (event: { sheetId: string; name: string }) => void;
  /**
   * Called after a sheet move is successfully applied to the metadata store.
   *
   * The desktop shell can use this to persist the new sheet order to the backend.
   */
  onSheetMoved?: (event: { sheetId: string; toIndex: number }) => Promise<void> | void;
  /**
   * Optional toast/error surface (used by the desktop shell).
   */
  onError?: (message: string) => void;
};

export function SheetTabStrip({
  store,
  activeSheetId,
  onActivateSheet,
  onAddSheet,
  onPersistSheetRename,
  onPersistSheetDelete,
  onSheetsReordered,
  onSheetRenamed,
  onSheetDeleted,
  onSheetMoved,
  onError,
}: Props) {
  const [sheets, setSheets] = useState<SheetMeta[]>(() => store.listAll());
  const activeSheetIdRef = useRef(activeSheetId);

  useEffect(() => {
    activeSheetIdRef.current = activeSheetId;
  }, [activeSheetId]);

  useEffect(() => {
    setSheets(store.listAll());
    return store.subscribe(() => {
      setSheets(store.listAll());
    });
  }, [store]);

  const visibleSheets = useMemo(() => sheets.filter((s) => s.visibility === "visible"), [sheets]);
  const [draggingSheetId, setDraggingSheetId] = useState<string | null>(null);
  const [dropIndicator, setDropIndicator] = useState<{ targetSheetId: string; position: "before" | "after" } | null>(null);

  const containerRef = useRef<HTMLDivElement | null>(null);
  const autoScrollRef = useRef<{ raf: number | null; direction: -1 | 0 | 1 }>({ raf: null, direction: 0 });
  const activeTabRef = useRef<HTMLButtonElement | null>(null);

  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState("");
  const [renameError, setRenameError] = useState<string | null>(null);
  const renameInputRef = useRef<HTMLInputElement>(null!);
  const renameCommitRef = useRef<Promise<boolean> | null>(null);
  const [canScroll, setCanScroll] = useState<{ left: boolean; right: boolean }>({ left: false, right: false });

  const lastContextMenuTabRef = useRef<HTMLButtonElement | null>(null);
  const tabContextMenu = useMemo(
    () =>
      new ContextMenu({
        testId: "sheet-tab-context-menu",
        onClose: () => {
          // Restore focus so keyboard users don't "fall off" the tab strip after dismissing the menu.
          if (lastContextMenuTabRef.current?.isConnected) {
            lastContextMenuTabRef.current.focus({ preventScroll: true });
          }
        },
      }),
    [],
  );

  useEffect(() => {
    return () => {
      tabContextMenu.close();
    };
  }, [tabContextMenu]);

  const stopAutoScroll = () => {
    const raf = autoScrollRef.current.raf;
    if (raf != null) cancelAnimationFrame(raf);
    autoScrollRef.current.raf = null;
    autoScrollRef.current.direction = 0;
  };

  const maybeAutoScroll = (clientX: number) => {
    const el = containerRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const threshold = 32;

    let direction: -1 | 0 | 1 = 0;
    if (clientX < rect.left + threshold) direction = -1;
    else if (clientX > rect.right - threshold) direction = 1;

    autoScrollRef.current.direction = direction;
    if (direction === 0) {
      stopAutoScroll();
      return;
    }

    if (autoScrollRef.current.raf != null) return;

    const tick = () => {
      const container = containerRef.current;
      if (!container) {
        stopAutoScroll();
        return;
      }
      const dir = autoScrollRef.current.direction;
      if (dir === 0) {
        stopAutoScroll();
        return;
      }
      container.scrollLeft += dir * 8;
      autoScrollRef.current.raf = requestAnimationFrame(tick);
    };
    autoScrollRef.current.raf = requestAnimationFrame(tick);
  };

  useEffect(() => {
    return () => {
      stopAutoScroll();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const clearDragIndicators = () => {
    setDropIndicator(null);
    setDraggingSheetId(null);
  };

  const commitRename = (sheetId: string): Promise<boolean> => {
    if (renameCommitRef.current) return renameCommitRef.current;

    const promise = (async () => {
      const oldName = store.getName(sheetId) ?? "";

      let normalized: string;
      try {
        normalized = validateSheetName(draftName, {
          sheets: store.listAll(),
          ignoreId: sheetId,
        });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        setRenameError(message);
        onError?.(message);
        // Keep editing: refocus the input.
        requestAnimationFrame(() => renameInputRef.current?.focus());
        return false;
      }

      // Treat no-op renames as a simple exit from rename mode.
      if (oldName && normalized === oldName) {
        setRenameError(null);
        setEditingSheetId(null);
        return true;
      }

      try {
        await onPersistSheetRename?.(sheetId, normalized);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        setRenameError(message);
        onError?.(message);
        requestAnimationFrame(() => renameInputRef.current?.focus());
        return false;
      }

      try {
        store.rename(sheetId, normalized);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        setRenameError(message);
        onError?.(message);
        // Keep editing: refocus the input.
        requestAnimationFrame(() => renameInputRef.current?.focus());
        return false;
      }

      const newName = store.getName(sheetId) ?? normalized;
      setRenameError(null);
      setEditingSheetId(null);
      if (oldName && newName && oldName !== newName) {
        try {
          onSheetRenamed?.({ sheetId, oldName, newName });
        } catch (err) {
          onError?.(err instanceof Error ? err.message : String(err));
        }
      }
      return true;
    })().finally(() => {
      renameCommitRef.current = null;
    });

    renameCommitRef.current = promise;
    return promise;
  };

  const moveSheet = (sheetId: string, dropTarget: Parameters<typeof computeWorkbookSheetMoveIndex>[0]["dropTarget"]) => {
    const all = store.listAll();
    const fromIndex = all.findIndex((s) => s.id === sheetId);
    if (fromIndex < 0) return;
    const toIndex = computeWorkbookSheetMoveIndex({ sheets: all, fromSheetId: sheetId, dropTarget });
    if (toIndex == null) return;
    if (toIndex === fromIndex) return;
    try {
      store.move(sheetId, toIndex);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }
    if (onSheetMoved) {
      void Promise.resolve()
        .then(() => onSheetMoved({ sheetId, toIndex }))
        .catch((err) => {
          onError?.(err instanceof Error ? err.message : String(err));
        });
    }
    onSheetsReordered?.();
  };

  const activateSheetWithRenameGuard = async (sheetId: string) => {
    if (editingSheetId && editingSheetId !== sheetId) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }
    onActivateSheet(sheetId);
  };

  const beginRenameWithGuard = async (sheet: SheetMeta) => {
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }

    setEditingSheetId(sheet.id);
    setDraftName(sheet.name);
    setRenameError(null);
  };

  const openSheetPicker = useCallback(async () => {
    // Match the "Add sheet" behavior: if the user is mid-rename and the rename is invalid,
    // keep them in the rename flow instead of navigating away.
    if (editingSheetId) {
      const ok = commitRename(editingSheetId);
      if (!ok) return;
    }

    const selected = await showQuickPick(
      visibleSheets.map((sheet) => ({
        label: sheet.name,
        value: sheet.id,
      })),
      { placeHolder: "Sheets" },
    );

    if (!selected) return;
    activateSheetWithRenameGuard(selected);
  }, [activateSheetWithRenameGuard, commitRename, editingSheetId, visibleSheets]);

  const openSheetTabContextMenu = (sheetId: string, anchor: { x: number; y: number }) => {
    const sheet = store.getById(sheetId);
    if (!sheet) return;

    const allSheets = store.listAll();
    // Only allow unhiding "hidden" sheets. Excel does not offer UI affordances for
    // "veryHidden" sheets (those are typically VBA-only), so keep them out of the menu.
    const hiddenSheets = allSheets.filter((s) => s.visibility === "hidden");

    // Prevent deleting/hiding the last visible sheet.
    const canDelete = visibleSheets.length > 1;
    const canHide = sheet.visibility === "visible" && visibleSheets.length > 1;

    const items: ContextMenuItem[] = [
      {
        type: "item",
        label: "Rename",
        onSelect: () => {
          void beginRenameWithGuard(sheet);
        },
      },
      {
        type: "item",
        label: "Hide",
        enabled: canHide,
        onSelect: () => {
          const wasActive = sheet.id === activeSheetIdRef.current;
          let nextActiveId: string | null = null;
          if (wasActive) {
            const idx = visibleSheets.findIndex((s) => s.id === sheet.id);
            nextActiveId = idx === -1 ? null : (visibleSheets[idx + 1]?.id ?? visibleSheets[idx - 1]?.id ?? null);
          }

          try {
            store.hide(sheet.id);
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            onError?.(message);
            return;
          }

          if (wasActive) {
            const fallback = store.listVisible().at(0)?.id ?? null;
            const next = nextActiveId ?? fallback;
            if (next && next !== sheet.id) {
              onActivateSheet(next);
            }
            return;
          }
          onActivateSheet(activeSheetIdRef.current);
        },
      },
      {
        type: "submenu",
        label: "Unhide…",
        enabled: hiddenSheets.length > 0,
        items: hiddenSheets.map((hidden) => ({
          type: "item" as const,
          label: hidden.name,
          onSelect: () => {
            try {
              store.unhide(hidden.id);
            } catch (err) {
              const message = err instanceof Error ? err.message : String(err);
              onError?.(message);
              return;
            }
            onActivateSheet(activeSheetIdRef.current);
          },
        })),
      },
      {
        type: "submenu",
        label: "Tab Color",
        items: [
          {
            type: "item",
            label: "No Color",
            onSelect: () => {
              try {
                store.setTabColor(sheet.id, undefined);
              } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                onError?.(message);
              }
            },
          },
          { type: "separator" },
          ...SHEET_TAB_COLOR_PALETTE.map((entry) => {
            const rgb = resolveCssVar(entry.token, { fallback: entry.fallbackCss });
            return {
              type: "item" as const,
              label: entry.label,
              leading: { type: "swatch" as const, token: entry.token },
              onSelect: () => {
                try {
                  store.setTabColor(sheet.id, { rgb });
                } catch (err) {
                  const message = err instanceof Error ? err.message : String(err);
                  onError?.(message);
                }
              },
            };
          }),
        ],
      },
    ];

    items.push({ type: "separator" });
    items.push({
      type: "item",
      label: "Delete",
      enabled: canDelete,
      onSelect: () => {
        void deleteSheet(sheet);
      },
    });

    tabContextMenu.open({ x: anchor.x, y: anchor.y, items });
  };

  const deleteSheet = async (sheet: SheetMeta): Promise<void> => {
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }

    if (editingSheetId === sheet.id) return;

    let ok = false;
    try {
      ok = await nativeDialogs.confirm(`Delete sheet "${sheet.name}"?`);
    } catch {
      ok = false;
    }
    if (!ok) return;

    const deletedName = sheet.name;

    try {
      await onPersistSheetDelete?.(sheet.id);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }

    try {
      store.remove(sheet.id);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }

    if (sheet.id === activeSheetIdRef.current) {
      const next = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
      if (next && next !== sheet.id) {
        onActivateSheet(next);
      }
    } else {
      // If we deleted a non-active sheet, re-focus the current sheet surface so the
      // user doesn't lose keyboard focus (especially after keyboard-invoked deletes).
      onActivateSheet(activeSheetIdRef.current);
    }

    try {
      onSheetDeleted?.({ sheetId: sheet.id, name: deletedName });
    } catch (err) {
      onError?.(err instanceof Error ? err.message : String(err));
    }
  };

  const updateScrollButtons = useCallback(() => {
    const el = containerRef.current;
    if (!el) {
      setCanScroll({ left: false, right: false });
      return;
    }
    const maxScrollLeft = el.scrollWidth - el.clientWidth;
    setCanScroll({
      left: el.scrollLeft > 0,
      right: el.scrollLeft < maxScrollLeft - 1,
    });
  }, []);

  const scrollTabsBy = (delta: number) => {
    const el = containerRef.current;
    if (!el) return;
    el.scrollBy({ left: delta, behavior: "smooth" });
  };

  const isSheetDrag = (dt: DataTransfer): boolean => {
    // We set both the custom type and `text/plain` on dragStart. Some environments
    // can be finicky about exposing custom MIME types during drag operations.
    return dt.types.includes("text/sheet-id") || dt.types.includes("text/plain");
  };

  useEffect(() => {
    updateScrollButtons();
  }, [updateScrollButtons, visibleSheets.length]);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    updateScrollButtons();
    const onScroll = () => updateScrollButtons();
    el.addEventListener("scroll", onScroll, { passive: true });
    window.addEventListener("resize", onScroll);
    return () => {
      el.removeEventListener("scroll", onScroll);
      window.removeEventListener("resize", onScroll);
    };
  }, [updateScrollButtons]);

  useEffect(() => {
    // Keep the active tab visible when switching sheets via keyboard or programmatically.
    if (editingSheetId) return;
    activeTabRef.current?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [activeSheetId, editingSheetId, visibleSheets.length]);

  return (
    <>
      <div className="sheet-nav">
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs left"
          tabIndex={-1}
          onClick={() => scrollTabsBy(-120)}
          disabled={!canScroll.left}
        >
          ‹
        </button>
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs right"
          tabIndex={-1}
          onClick={() => scrollTabsBy(120)}
          disabled={!canScroll.right}
        >
          ›
        </button>
      </div>

      <div
        className="sheet-tabs"
        ref={containerRef}
        role="tablist"
        aria-label="Sheets"
        aria-orientation="horizontal"
        data-dragging={draggingSheetId ? "true" : undefined}
        onKeyDown={(e) => {
          if (e.defaultPrevented) return;

          const target = e.target as HTMLElement | null;
          if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;

          const tabs = containerRef.current
            ? Array.from(containerRef.current.querySelectorAll<HTMLButtonElement>('button[role="tab"]'))
            : [];

          if (target instanceof HTMLButtonElement && target.getAttribute("role") === "tab" && tabs.length > 0) {
            const idx = tabs.indexOf(target);
            if (idx !== -1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
              const focusTab = (nextIdx: number) => {
                const next = tabs[nextIdx];
                if (!next) return;
                next.focus();
                next.scrollIntoView({ block: "nearest", inline: "nearest" });
              };

              const isContextMenuKey = (e.shiftKey && e.key === "F10") || e.key === "ContextMenu" || e.code === "ContextMenu";
              if (isContextMenuKey) {
                e.preventDefault();
                e.stopPropagation();
                const sheetId = target.dataset.sheetId;
                if (!sheetId) return;
                const activeEditingId = editingSheetId;
                void (async () => {
                  if (activeEditingId && activeEditingId !== sheetId) {
                    const ok = await commitRename(activeEditingId);
                    if (!ok) return;
                  }
                  lastContextMenuTabRef.current = target;
                  const rect = target.getBoundingClientRect();
                  openSheetTabContextMenu(sheetId, { x: rect.left + rect.width / 2, y: rect.bottom });
                })();
                return;
              }

              if (e.key === "ArrowRight") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.min(tabs.length - 1, idx + 1));
                return;
              }
              if (e.key === "ArrowLeft") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.max(0, idx - 1));
                return;
              }
              if (e.key === "Home") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(0);
                return;
              }
              if (e.key === "End") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.max(0, tabs.length - 1));
                return;
              }

              // Enter/Space are handled by the <button> itself to activate the focused tab.
            }
          }

          const primary = e.ctrlKey || e.metaKey;
          if (!primary) return;
          if (e.shiftKey || e.altKey) return;
          if (e.key !== "PageUp" && e.key !== "PageDown") return;
          e.preventDefault();

          const idx = visibleSheets.findIndex((s) => s.id === activeSheetId);
          if (idx === -1) return;
          const delta = e.key === "PageUp" ? -1 : 1;
          const next = visibleSheets[(idx + delta + visibleSheets.length) % visibleSheets.length];
          if (next) void activateSheetWithRenameGuard(next.id);
        }}
        onDragOver={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          maybeAutoScroll(e.clientX);

          // Only apply an "end" indicator when dragging over the container itself (not over a tab),
          // otherwise the per-tab handler will compute before/after.
          if (e.target === e.currentTarget) {
            const last = visibleSheets.at(-1)?.id ?? null;
            if (last) setDropIndicator({ targetSheetId: last, position: "after" });
          }
        }}
        onDrop={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          stopAutoScroll();
          const fromId = e.dataTransfer.getData("text/sheet-id") || e.dataTransfer.getData("text/plain");
          if (!fromId) {
            clearDragIndicators();
            return;
          }
          // Dropping on the container inserts at the end of the visible list.
          moveSheet(fromId, { kind: "end" });
          clearDragIndicators();
        }}
        onDragLeave={(e) => {
          stopAutoScroll();
          // Avoid flickering the drop indicator when moving between child tabs.
          if (e.target === e.currentTarget) setDropIndicator(null);
        }}
      >
        {visibleSheets.map((sheet) => (
          <SheetTab
            key={sheet.id}
            sheet={sheet}
            active={sheet.id === activeSheetId}
            editing={editingSheetId === sheet.id}
            dragging={draggingSheetId === sheet.id}
            dropPosition={dropIndicator?.targetSheetId === sheet.id ? dropIndicator.position : null}
            draftName={draftName}
            renameError={editingSheetId === sheet.id ? renameError : null}
            renameInputRef={renameInputRef}
            tabRef={sheet.id === activeSheetId ? activeTabRef : undefined}
            onActivate={() => void activateSheetWithRenameGuard(sheet.id)}
            onBeginRename={() => void beginRenameWithGuard(sheet)}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              const activeEditingId = editingSheetId;
              const anchor = { x: e.clientX, y: e.clientY };
              const target = e.currentTarget;
              void (async () => {
                if (activeEditingId && activeEditingId !== sheet.id) {
                  const ok = await commitRename(activeEditingId);
                  if (!ok) return;
                }
                lastContextMenuTabRef.current = target;
                openSheetTabContextMenu(sheet.id, anchor);
              })();
            }}
            onCommitRename={() => {
              void commitRename(sheet.id);
            }}
            onCancelRename={() => {
              setEditingSheetId(null);
              setRenameError(null);
            }}
            onDraftNameChange={setDraftName}
            onDragStart={() => {
              stopAutoScroll();
              setDraggingSheetId(sheet.id);
              setDropIndicator(null);
            }}
            onDragEnd={() => {
              stopAutoScroll();
              clearDragIndicators();
            }}
            onDragOverTab={(e) => {
              if (!isSheetDrag(e.dataTransfer)) return;
              if (draggingSheetId && draggingSheetId === sheet.id) {
                setDropIndicator(null);
                return;
              }

              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const shouldInsertAfter = e.clientX > rect.left + rect.width / 2;
              setDropIndicator({ targetSheetId: sheet.id, position: shouldInsertAfter ? "after" : "before" });
            }}
            onDropOnTab={(e) => {
              stopAutoScroll();
              clearDragIndicators();
              const fromId = e.dataTransfer.getData("text/sheet-id") || e.dataTransfer.getData("text/plain");
              if (!fromId || fromId === sheet.id) return;

              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const shouldInsertAfter = e.clientX > rect.left + rect.width / 2;
              moveSheet(fromId, {
                kind: shouldInsertAfter ? "after" : "before",
                targetSheetId: sheet.id,
              });
            }}
          />
        ))}
      </div>

      <button
        type="button"
        className="sheet-add"
        data-testid="sheet-add"
        onClick={() => {
          void (async () => {
            if (editingSheetId) {
              const ok = await commitRename(editingSheetId);
              if (!ok) return;
            }
            await onAddSheet();
          })();
        }}
        aria-label="Add sheet"
      >
        +
      </button>

      <button
        type="button"
        className="sheet-overflow"
        data-testid="sheet-overflow"
        aria-label="Show sheet list"
        onClick={() => void openSheetPicker()}
      >
        ⋯
      </button>
    </>
  );
}

function SheetTab(props: {
  sheet: SheetMeta;
  active: boolean;
  editing: boolean;
  dragging: boolean;
  dropPosition: "before" | "after" | null;
  draftName: string;
  renameError: string | null;
  renameInputRef: React.RefObject<HTMLInputElement>;
  tabRef?: React.Ref<HTMLButtonElement>;
  onActivate: () => void;
  onBeginRename: () => void;
  onContextMenu: (e: React.MouseEvent<HTMLButtonElement>) => void;
  onCommitRename: () => void;
  onCancelRename: () => void;
  onDraftNameChange: (name: string) => void;
  onDragStart: () => void;
  onDragEnd: () => void;
  onDragOverTab: (e: React.DragEvent<HTMLButtonElement>) => void;
  onDropOnTab: (e: React.DragEvent<HTMLButtonElement>) => void;
}) {
  const { sheet, active, editing, draftName, renameError } = props;
  const cancelBlurCommitRef = useRef(false);
  const tabColorCss = !editing ? (normalizeExcelColorToCss(sheet.tabColor?.rgb) ?? null) : null;
  const ariaLabel = sheet.visibility === "visible" ? sheet.name : `${sheet.name} (${sheet.visibility})`;

  return (
    <button
      type="button"
      className="sheet-tab"
      role="tab"
      aria-selected={active}
      aria-label={ariaLabel}
      tabIndex={active ? 0 : -1}
      data-testid={`sheet-tab-${sheet.id}`}
      data-sheet-id={sheet.id}
      data-active={active ? "true" : "false"}
      data-tab-color={tabColorCss ?? undefined}
      data-dragging={props.dragging ? "true" : undefined}
      data-drop-position={props.dropPosition ?? undefined}
      draggable={!editing}
      ref={props.tabRef}
      onClick={() => {
        if (!editing) props.onActivate();
      }}
      onDoubleClick={() => {
        if (!editing) props.onBeginRename();
      }}
      onContextMenu={(e) => {
        if (editing) return;
        props.onContextMenu(e);
      }}
      onDragStart={(e) => {
        props.onDragStart();
        e.dataTransfer.setData("text/sheet-id", sheet.id);
        e.dataTransfer.setData("text/plain", sheet.id);
        e.dataTransfer.effectAllowed = "move";
      }}
      onDragOver={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id") && !e.dataTransfer.types.includes("text/plain")) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        props.onDragOverTab(e);
      }}
      onDrop={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id") && !e.dataTransfer.types.includes("text/plain")) return;
        e.preventDefault();
        e.stopPropagation();
        props.onDropOnTab(e);
      }}
      onDragEnd={() => props.onDragEnd()}
    >
      {editing ? (
        <input
          ref={props.renameInputRef}
          className="sheet-tab__input"
          value={draftName}
          autoFocus
          aria-invalid={renameError ? true : undefined}
          title={renameError ?? undefined}
          onClick={(e) => e.stopPropagation()}
          onChange={(e) => props.onDraftNameChange(e.target.value)}
          onFocus={(e) => e.currentTarget.select()}
          onBlur={() => {
            if (cancelBlurCommitRef.current) {
              cancelBlurCommitRef.current = false;
              return;
            }
            props.onCommitRename();
          }}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === "Enter") {
              e.preventDefault();
              props.onCommitRename();
            }
            if (e.key === "Escape") {
              e.preventDefault();
              cancelBlurCommitRef.current = true;
              props.onCancelRename();
            }
          }}
        />
      ) : (
        <>
          <span className="sheet-tab__name">{sheet.name}</span>
          {tabColorCss ? <span className="sheet-tab__color" style={{ background: tabColorCss }} /> : null}
        </>
      )}
    </button>
  );
}
