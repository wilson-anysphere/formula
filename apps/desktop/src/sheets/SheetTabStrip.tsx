import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { ContextMenu, type ContextMenuItem } from "../menus/contextMenu";
import * as nativeDialogs from "../tauri/nativeDialogs";
import type { SheetMeta, TabColor, WorkbookSheetStore } from "./workbookSheetStore";
import { computeWorkbookSheetMoveIndex } from "./sheetReorder";

type Props = {
  store: WorkbookSheetStore;
  activeSheetId: string;
  onActivateSheet: (sheetId: string) => void;
  onAddSheet: () => Promise<void> | void;
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
   * Optional toast/error surface (used by the desktop shell).
   */
  onError?: (message: string) => void;
};

export function SheetTabStrip({
  store,
  activeSheetId,
  onActivateSheet,
  onAddSheet,
  onSheetRenamed,
  onSheetDeleted,
  onError,
}: Props) {
  const [sheets, setSheets] = useState<SheetMeta[]>(() => store.listAll());

  useEffect(() => {
    setSheets(store.listAll());
    return store.subscribe(() => {
      setSheets(store.listAll());
    });
  }, [store]);

  const visibleSheets = useMemo(() => sheets.filter((s) => s.visibility === "visible"), [sheets]);

  const containerRef = useRef<HTMLDivElement | null>(null);
  const autoScrollRef = useRef<{ raf: number | null; direction: -1 | 0 | 1 }>({ raf: null, direction: 0 });
  const activeTabRef = useRef<HTMLButtonElement | null>(null);

  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState("");
  const [renameError, setRenameError] = useState<string | null>(null);
  const renameInputRef = useRef<HTMLInputElement>(null!);
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

  const commitRename = (sheetId: string): boolean => {
    const oldName = store.getName(sheetId) ?? "";
    try {
      store.rename(sheetId, draftName);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setRenameError(message);
      onError?.(message);
      // Keep editing: refocus the input.
      requestAnimationFrame(() => renameInputRef.current?.focus());
      return false;
    }
    const newName = store.getName(sheetId) ?? draftName.trim();
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
  };

  const moveSheet = (sheetId: string, dropTarget: Parameters<typeof computeWorkbookSheetMoveIndex>[0]["dropTarget"]) => {
    const all = store.listAll();
    const toIndex = computeWorkbookSheetMoveIndex({ sheets: all, fromSheetId: sheetId, dropTarget });
    if (toIndex == null) return;
    try {
      store.move(sheetId, toIndex);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
    }
  };

  const activateSheetWithRenameGuard = (sheetId: string) => {
    if (editingSheetId && editingSheetId !== sheetId) {
      const ok = commitRename(editingSheetId);
      if (!ok) return;
    }
    onActivateSheet(sheetId);
  };

  const beginRenameWithGuard = (sheet: SheetMeta) => {
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = commitRename(editingSheetId);
      if (!ok) return;
    }

    setEditingSheetId(sheet.id);
    setDraftName(sheet.name);
    setRenameError(null);
  };

  const openSheetTabContextMenu = (sheetId: string, anchor: { x: number; y: number }) => {
    const sheet = store.getById(sheetId);
    if (!sheet) return;

    // Prevent deleting the last visible sheet (we currently have no unhide UI, and
    // having zero visible sheets would leave the workbook unusable).
    const canDelete = visibleSheets.length > 1;

    const items: ContextMenuItem[] = [
      {
        type: "item",
        label: "Rename",
        onSelect: () => beginRenameWithGuard(sheet),
      },
      {
        type: "item",
        label: "Delete",
        enabled: canDelete,
        onSelect: () => {
          void deleteSheet(sheet);
        },
      },
    ];

    tabContextMenu.open({ x: anchor.x, y: anchor.y, items });
  };

  const deleteSheet = async (sheet: SheetMeta): Promise<void> => {
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = commitRename(editingSheetId);
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
      store.remove(sheet.id);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }

    if (sheet.id === activeSheetId) {
      const next = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
      if (next && next !== sheet.id) {
        onActivateSheet(next);
      }
    } else {
      // If we deleted a non-active sheet, re-focus the current sheet surface so the
      // user doesn't lose keyboard focus (especially after keyboard-invoked deletes).
      onActivateSheet(activeSheetId);
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
                if (editingSheetId && editingSheetId !== sheetId) {
                  const ok = commitRename(editingSheetId);
                  if (!ok) return;
                }
                lastContextMenuTabRef.current = target;
                const rect = target.getBoundingClientRect();
                openSheetTabContextMenu(sheetId, { x: rect.left + rect.width / 2, y: rect.bottom });
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
          if (next) activateSheetWithRenameGuard(next.id);
        }}
        onDragOver={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          maybeAutoScroll(e.clientX);
        }}
        onDrop={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          stopAutoScroll();
          const fromId = e.dataTransfer.getData("text/sheet-id") || e.dataTransfer.getData("text/plain");
          if (!fromId) return;
          // Dropping on the container inserts at the end of the visible list.
          moveSheet(fromId, { kind: "end" });
        }}
        onDragLeave={() => {
          stopAutoScroll();
        }}
      >
        {visibleSheets.map((sheet) => (
          <SheetTab
            key={sheet.id}
            sheet={sheet}
            active={sheet.id === activeSheetId}
            editing={editingSheetId === sheet.id}
            draftName={draftName}
            renameError={editingSheetId === sheet.id ? renameError : null}
            renameInputRef={renameInputRef}
            tabRef={sheet.id === activeSheetId ? activeTabRef : undefined}
            onActivate={() => activateSheetWithRenameGuard(sheet.id)}
            onBeginRename={() => beginRenameWithGuard(sheet)}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              if (editingSheetId && editingSheetId !== sheet.id) {
                const ok = commitRename(editingSheetId);
                if (!ok) return;
              }
              lastContextMenuTabRef.current = e.currentTarget;
              openSheetTabContextMenu(sheet.id, { x: e.clientX, y: e.clientY });
            }}
            onCommitRename={() => commitRename(sheet.id)}
            onCancelRename={() => {
              setEditingSheetId(null);
              setRenameError(null);
            }}
            onDraftNameChange={setDraftName}
            onDragStart={() => {
              stopAutoScroll();
            }}
            onDragEnd={() => {
              stopAutoScroll();
            }}
            onDropOnTab={(e) => {
              stopAutoScroll();
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
          if (editingSheetId) {
            const ok = commitRename(editingSheetId);
            if (!ok) return;
          }
          void onAddSheet();
        }}
        aria-label="Add sheet"
      >
        +
      </button>
    </>
  );
}

function SheetTab(props: {
  sheet: SheetMeta;
  active: boolean;
  editing: boolean;
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
  onDropOnTab: (e: React.DragEvent<HTMLButtonElement>) => void;
}) {
  const { sheet, active, editing, draftName, renameError } = props;
  const cancelBlurCommitRef = useRef(false);
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
          {sheet.tabColor?.rgb ? <TabColorUnderline tabColor={sheet.tabColor} /> : null}
        </>
      )}
    </button>
  );
}

function TabColorUnderline({ tabColor }: { tabColor: TabColor }) {
  const rgb = tabColor.rgb;
  if (!rgb) return null;
  const css = tabColorRgbToCss(rgb);
  return <span className="sheet-tab__color" style={{ background: css }} />;
}

function tabColorRgbToCss(raw: string): string {
  const rgb = String(raw ?? "").trim();
  if (!rgb) return "transparent";
  if (/^#[0-9A-Fa-f]{6}$/.test(rgb)) return rgb;
  if (/^[0-9A-Fa-f]{6}$/.test(rgb)) return `#${rgb}`;
  if (/^#[0-9A-Fa-f]{8}$/.test(rgb)) return argbToCssHsl(rgb.slice(1));
  if (/^[0-9A-Fa-f]{8}$/.test(rgb)) return argbToCssHsl(rgb);
  // Best-effort fallback (handles named colors, rgb/rgba strings, etc).
  return rgb;
}

function argbToCssHsl(argb: string): string {
  if (!/^([0-9A-Fa-f]{8})$/.test(argb)) return "transparent";
  const alpha = parseInt(argb.slice(0, 2), 16) / 255;
  const r = parseInt(argb.slice(2, 4), 16);
  const g = parseInt(argb.slice(4, 6), 16);
  const b = parseInt(argb.slice(6, 8), 16);

  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const delta = max - min;
  const light = (max + min) / 2;

  let hue = 0;
  let sat = 0;

  if (delta !== 0) {
    sat = delta / (1 - Math.abs(2 * light - 1));
    switch (max) {
      case rn:
        hue = ((gn - bn) / delta + (gn < bn ? 6 : 0)) * 60;
        break;
      case gn:
        hue = ((bn - rn) / delta + 2) * 60;
        break;
      default:
        hue = ((rn - gn) / delta + 4) * 60;
        break;
    }
  }

  const h = Math.round(hue);
  const s = Math.round(sat * 100);
  const l = Math.round(light * 100);
  const a = Math.max(0, Math.min(1, alpha));

  return a < 1 ? `hsla(${h}, ${s}%, ${l}%, ${a.toFixed(3)})` : `hsl(${h}, ${s}%, ${l}%)`;
}
