import React, { useEffect, useMemo, useRef, useState } from "react";

import type { SheetMeta, TabColor, WorkbookSheetStore } from "./workbookSheetStore";

type Props = {
  store: WorkbookSheetStore;
  activeSheetId: string;
  onActivateSheet: (sheetId: string) => void;
  onAddSheet: () => Promise<void> | void;
  /**
   * Optional toast/error surface (used by the desktop shell).
   */
  onError?: (message: string) => void;
};

export function SheetTabStrip({ store, activeSheetId, onActivateSheet, onAddSheet, onError }: Props) {
  const [sheets, setSheets] = useState<SheetMeta[]>(() => store.listAll());

  useEffect(() => {
    setSheets(store.listAll());
    return store.subscribe(() => {
      setSheets(store.listAll());
    });
  }, [store]);

  const visibleSheets = useMemo(() => sheets.filter((s) => s.visibility === "visible"), [sheets]);

  const containerRef = useRef<HTMLDivElement | null>(null);
  const dragSheetIdRef = useRef<string | null>(null);
  const autoScrollRef = useRef<{ raf: number | null; direction: -1 | 0 | 1 }>({ raf: null, direction: 0 });

  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [draftName, setDraftName] = useState("");
  const [renameError, setRenameError] = useState<string | null>(null);
  const renameInputRef = useRef<HTMLInputElement>(null!);

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

  const commitRename = (sheetId: string) => {
    try {
      store.rename(sheetId, draftName);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setRenameError(message);
      onError?.(message);
      // Keep editing: refocus the input.
      requestAnimationFrame(() => renameInputRef.current?.focus());
      return;
    }
    setRenameError(null);
    setEditingSheetId(null);
  };

  const moveVisibleSheet = (sheetId: string, targetVisibleIndex: number) => {
    // Preserve hidden sheets by only reordering visible slots.
    const all = store.listAll();
    const visible = all.filter((s) => s.visibility === "visible").map((s) => s.id);
    if (visible.length <= 1) return;
    const fromIndex = visible.findIndex((id) => id === sheetId);
    if (fromIndex < 0) return;

    const rawTarget = Math.max(0, Math.min(Math.floor(targetVisibleIndex), visible.length));
    const insertIndex = rawTarget > fromIndex ? rawTarget - 1 : rawTarget;
    if (insertIndex === fromIndex) return;

    const nextVisible = visible.slice();
    nextVisible.splice(fromIndex, 1);
    nextVisible.splice(insertIndex, 0, sheetId);

    const desiredAll: string[] = [];
    let v = 0;
    for (const sheet of all) {
      if (sheet.visibility !== "visible") {
        desiredAll.push(sheet.id);
        continue;
      }
      desiredAll.push(nextVisible[v] ?? sheet.id);
      v += 1;
    }

    const currentIds = all.map((s) => s.id);
    const moveInArray = (ids: string[], from: number, to: number) => {
      const [item] = ids.splice(from, 1);
      if (!item) return;
      ids.splice(to, 0, item);
    };

    for (let i = 0; i < desiredAll.length; i += 1) {
      const desiredId = desiredAll[i];
      if (!desiredId) continue;
      if (currentIds[i] === desiredId) continue;
      const from = currentIds.indexOf(desiredId);
      if (from < 0) continue;
      try {
        store.move(desiredId, i);
      } catch {
        return;
      }
      moveInArray(currentIds, from, i);
    }
  };

  const scrollTabsBy = (delta: number) => {
    const el = containerRef.current;
    if (!el) return;
    el.scrollBy({ left: delta, behavior: "smooth" });
  };

  return (
    <>
      <div className="sheet-nav">
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs left"
          onClick={() => scrollTabsBy(-120)}
        >
          ‹
        </button>
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs right"
          onClick={() => scrollTabsBy(120)}
        >
          ›
        </button>
      </div>

      <div
        className="sheet-tabs"
        ref={containerRef}
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.defaultPrevented) return;
          const primary = e.ctrlKey || e.metaKey;
          if (!primary) return;
          if (e.shiftKey || e.altKey) return;
          if (e.key !== "PageUp" && e.key !== "PageDown") return;

          const target = e.target as HTMLElement | null;
          if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;
          e.preventDefault();

          const idx = visibleSheets.findIndex((s) => s.id === activeSheetId);
          if (idx === -1) return;
          const delta = e.key === "PageUp" ? -1 : 1;
          const next = visibleSheets[(idx + delta + visibleSheets.length) % visibleSheets.length];
          if (next) onActivateSheet(next.id);
        }}
        onDragOver={(e) => {
          if (!e.dataTransfer.types.includes("text/sheet-id")) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          maybeAutoScroll(e.clientX);
        }}
        onDrop={(e) => {
          if (!e.dataTransfer.types.includes("text/sheet-id")) return;
          e.preventDefault();
          stopAutoScroll();
          const fromId = e.dataTransfer.getData("text/sheet-id");
          if (!fromId) return;
          // Dropping on the container inserts at the end of the visible list.
          moveVisibleSheet(fromId, visibleSheets.length);
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
            onActivate={() => onActivateSheet(sheet.id)}
            onBeginRename={() => {
              setEditingSheetId(sheet.id);
              setDraftName(sheet.name);
              setRenameError(null);
            }}
            onCommitRename={() => commitRename(sheet.id)}
            onCancelRename={() => {
              setEditingSheetId(null);
              setRenameError(null);
            }}
            onDraftNameChange={setDraftName}
            onDragStart={() => {
              dragSheetIdRef.current = sheet.id;
              stopAutoScroll();
            }}
            onDragEnd={() => {
              dragSheetIdRef.current = null;
              stopAutoScroll();
            }}
            onDropOnTab={(e) => {
              stopAutoScroll();
              const fromId = e.dataTransfer.getData("text/sheet-id");
              if (!fromId || fromId === sheet.id) return;

              const targetIndex = visibleSheets.findIndex((s) => s.id === sheet.id);
              if (targetIndex < 0) return;

              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const shouldInsertAfter = e.clientX > rect.left + rect.width / 2;
              const nextIndex = targetIndex + (shouldInsertAfter ? 1 : 0);

              moveVisibleSheet(fromId, nextIndex);
            }}
          />
        ))}
      </div>

      <button
        type="button"
        className="sheet-add"
        data-testid="sheet-add"
        onClick={() => void onAddSheet()}
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
  onActivate: () => void;
  onBeginRename: () => void;
  onCommitRename: () => void;
  onCancelRename: () => void;
  onDraftNameChange: (name: string) => void;
  onDragStart: () => void;
  onDragEnd: () => void;
  onDropOnTab: (e: React.DragEvent<HTMLButtonElement>) => void;
}) {
  const { sheet, active, editing, draftName, renameError } = props;

  return (
    <button
      type="button"
      className="sheet-tab"
      data-testid={`sheet-tab-${sheet.id}`}
      data-sheet-id={sheet.id}
      data-active={active ? "true" : "false"}
      draggable={!editing}
      onClick={() => {
        if (!editing) props.onActivate();
      }}
      onDoubleClick={() => {
        if (!editing) props.onBeginRename();
      }}
      onDragStart={(e) => {
        props.onDragStart();
        e.dataTransfer.setData("text/sheet-id", sheet.id);
        e.dataTransfer.effectAllowed = "move";
      }}
      onDragOver={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id")) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
      }}
      onDrop={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id")) return;
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
          onBlur={() => props.onCommitRename()}
          onKeyDown={(e) => {
            e.stopPropagation();
            if (e.key === "Enter") {
              e.preventDefault();
              props.onCommitRename();
            }
            if (e.key === "Escape") {
              e.preventDefault();
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
  // Best-effort fallback (handles named colors, rgb(...), etc).
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
