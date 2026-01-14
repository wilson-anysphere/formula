import React, { useCallback, useEffect, useRef, useState } from "react";

import type { DrawingObject, DrawingObjectId } from "../../drawings/types";
import { isChartStoreDrawingId } from "../../charts/chartDrawingAdapter";

import { ChartIcon } from "../../ui/icons/ChartIcon";
import { BringForwardIcon } from "../../ui/icons/BringForwardIcon";
import { PictureIcon } from "../../ui/icons/PictureIcon";
import { SendBackwardIcon } from "../../ui/icons/SendBackwardIcon";
import { ShapesIcon } from "../../ui/icons/ShapesIcon";
import { TrashIcon } from "../../ui/icons/TrashIcon";

type SelectionPaneApp = {
  listDrawingsForSheet(sheetId?: string): DrawingObject[];
  subscribeDrawings(listener: () => void): () => void;
  getSelectedDrawingId(): DrawingObjectId | null;
  subscribeDrawingSelection(listener: (id: DrawingObjectId | null) => void): () => void;
  selectDrawingById(id: DrawingObjectId | null): void;
  deleteDrawingById?(id: DrawingObjectId): void;
  bringSelectedDrawingForward?(): void;
  sendSelectedDrawingBackward?(): void;
  getCurrentSheetId?(): string;
  isReadOnly?(): boolean;
  isEditing?(): boolean;
  onEditStateChange?(listener: (isEditing: boolean) => void): () => void;
  focus?(): void;
};

function DrawingKindIcon({ kind }: { kind: DrawingObject["kind"]["type"] }) {
  switch (kind) {
    case "image":
      return <PictureIcon size={16} />;
    case "shape":
      return <ShapesIcon size={16} />;
    case "chart":
      return <ChartIcon size={16} />;
    default:
      return <ShapesIcon size={16} />;
  }
}

type SelectionPaneItem = { obj: DrawingObject; label: string };

type StoredLabel = { kind: DrawingObject["kind"]["type"]; label: string; auto: boolean };

type SheetLabelState = {
  counters: Record<string, number>;
  labels: Map<DrawingObjectId, StoredLabel>;
};

export function SelectionPanePanel({ app }: { app: SelectionPaneApp }) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const labelsBySheetRef = useRef<Map<string, SheetLabelState>>(new Map());

  const [isReadOnly, setIsReadOnly] = useState<boolean>(() => {
    if (typeof app.isReadOnly !== "function") return false;
    try {
      return Boolean(app.isReadOnly());
    } catch {
      return false;
    }
  });

  const [isEditing, setIsEditing] = useState<boolean>(() => {
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    if (typeof globalEditing === "boolean") return globalEditing;
    if (typeof app.isEditing !== "function") return false;
    try {
      return Boolean(app.isEditing());
    } catch {
      return false;
    }
  });

  const actionsDisabled = isReadOnly || isEditing;

  useEffect(() => {
    if (typeof app.onEditStateChange !== "function") return;
    return app.onEditStateChange(() => {
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      if (typeof globalEditing === "boolean") {
        setIsEditing(globalEditing);
        return;
      }
      try {
        setIsEditing(Boolean(app.isEditing?.()));
      } catch {
        setIsEditing(false);
      }
    });
  }, [app]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onReadOnlyChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.readOnly === "boolean") {
        setIsReadOnly(detail.readOnly);
        return;
      }
      if (typeof app.isReadOnly !== "function") return;
      try {
        setIsReadOnly(Boolean(app.isReadOnly()));
      } catch {
        // ignore
      }
    };
    window.addEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
    return () => window.removeEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
  }, [app]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onEditingChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.isEditing === "boolean") {
        setIsEditing(detail.isEditing);
        return;
      }
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      if (typeof globalEditing === "boolean") {
        setIsEditing(globalEditing);
        return;
      }
      if (typeof app.isEditing !== "function") return;
      try {
        setIsEditing(Boolean(app.isEditing()));
      } catch {
        // ignore
      }
    };
    window.addEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
    return () => window.removeEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
  }, [app]);

  const computeItems = useCallback(
    (drawings: DrawingObject[]): SelectionPaneItem[] => {
      const sheetId = (() => {
        if (typeof app.getCurrentSheetId !== "function") return "__default";
        try {
          const id = String(app.getCurrentSheetId() ?? "").trim();
          return id || "__default";
        } catch {
          return "__default";
        }
      })();

      let sheetState = labelsBySheetRef.current.get(sheetId);
      if (!sheetState) {
        sheetState = { counters: Object.create(null) as Record<string, number>, labels: new Map() };
        labelsBySheetRef.current.set(sheetId, sheetState);
      }

      const seen = new Set<DrawingObjectId>();
      const nextItems: SelectionPaneItem[] = [];

      for (const obj of drawings) {
        seen.add(obj.id);

        const kind = obj.kind.type;
        const explicitLabel = (obj.kind as { label?: string }).label?.trim();
        if (explicitLabel) {
          sheetState.labels.set(obj.id, { kind, label: explicitLabel, auto: false });
          nextItems.push({ obj, label: explicitLabel });
          continue;
        }

        const existing = sheetState.labels.get(obj.id);
        if (existing && (!existing.auto || existing.kind === kind)) {
          nextItems.push({ obj, label: existing.label });
          continue;
        }

        const { counterKey, prefix, includeId } = (() => {
          // Excel uses "Picture" (not "Image") for inserted images.
          if (kind === "image") return { counterKey: "Picture", prefix: "Picture", includeId: false };
          if (kind === "shape") return { counterKey: "Shape", prefix: "Shape", includeId: false };
          if (kind === "chart") return { counterKey: "Chart", prefix: "Chart", includeId: false };
          const normalized = kind ? kind.slice(0, 1).toUpperCase() + kind.slice(1) : "Object";
          return { counterKey: normalized, prefix: normalized, includeId: true };
        })();

        const nextIndex = (sheetState.counters[counterKey] ?? 0) + 1;
        sheetState.counters[counterKey] = nextIndex;

        const label = includeId ? `${prefix} ${nextIndex} (id=${obj.id})` : `${prefix} ${nextIndex}`;
        sheetState.labels.set(obj.id, { kind, label, auto: true });
        nextItems.push({ obj, label });
      }

      // Prune labels for objects no longer present on this sheet to avoid leaking memory.
      for (const id of sheetState.labels.keys()) {
        if (!seen.has(id)) sheetState.labels.delete(id);
      }

      return nextItems;
    },
    [app],
  );

  const [items, setItems] = useState<SelectionPaneItem[]>(() => computeItems(app.listDrawingsForSheet()));
  const [selectedId, setSelectedId] = useState<DrawingObjectId | null>(() => app.getSelectedDrawingId());

  useEffect(() => {
    const onDrawings = () => {
      setItems(computeItems(app.listDrawingsForSheet()));
    };
    return app.subscribeDrawings(onDrawings);
  }, [app, computeItems]);

  useEffect(() => {
    const onSelection = (id: DrawingObjectId | null) => {
      setSelectedId(id);
    };
    return app.subscribeDrawingSelection(onSelection);
  }, [app]);

  const scrollItemIntoView = useCallback((id: DrawingObjectId) => {
    if (typeof document === "undefined") return;
    const root = rootRef.current;
    if (!root) return;
    try {
      const el = root.querySelector<HTMLElement>(`[data-testid="selection-pane-item-${id}"]`);
      el?.scrollIntoView?.({ block: "nearest" });
    } catch {
      // ignore
    }
  }, []);

  // When selection changes externally (e.g. clicking a drawing on the grid), keep the
  // selected row visible (Excel-like behavior) without stealing focus.
  useEffect(() => {
    if (selectedId == null) return;
    scrollItemIntoView(selectedId);
  }, [scrollItemIntoView, selectedId]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      if (items.length === 0) return;
      // For most interactions, only handle keys when the Selection Pane root itself is focused.
      // (If a per-row action button is focused, allow normal Tab navigation and avoid hijacking
      // arrow-key navigation, etc.)
      //
      // Exception: some keys should behave like global Selection Pane commands even when a child
      // element is focused (e.g. Delete should delete the selected object, not the active cell).
      const rootHasFocus = e.target === e.currentTarget;
      if (e.ctrlKey || e.metaKey || e.altKey) return;

      const currentIndex = selectedId == null ? -1 : items.findIndex(({ obj }) => obj.id === selectedId);

      const selectIndex = (nextIndex: number) => {
        if (nextIndex < 0 || nextIndex >= items.length) return;
        const nextId = items[nextIndex]!.obj.id;
        if (nextId === selectedId) return;
        e.preventDefault();
        e.stopPropagation();
        app.selectDrawingById(nextId);
        scrollItemIntoView(nextId);
      };

      switch (e.key) {
        case "ArrowDown":
          if (!rootHasFocus) return;
          selectIndex(currentIndex < 0 ? 0 : Math.min(currentIndex + 1, items.length - 1));
          return;
        case "ArrowUp":
          if (!rootHasFocus) return;
          selectIndex(currentIndex < 0 ? items.length - 1 : Math.max(currentIndex - 1, 0));
          return;
        case "Home":
          if (!rootHasFocus) return;
          selectIndex(0);
          return;
        case "End":
          if (!rootHasFocus) return;
          selectIndex(items.length - 1);
          return;
        case "Delete":
        case "Backspace": {
          if (selectedId == null) return;
          if (typeof app.deleteDrawingById !== "function") return;
          if (actionsDisabled) return;
          e.preventDefault();
          e.stopPropagation();
          app.deleteDrawingById(selectedId);
          return;
        }
        case "Escape": {
          e.preventDefault();
          e.stopPropagation();
          if (typeof app.focus === "function") {
            app.focus();
          }
          return;
        }
        default:
          return;
      }
    },
    [actionsDisabled, app, items, scrollItemIntoView, selectedId],
  );

  // When canvas charts are enabled (default), ChartStore charts render as drawing objects with a
  // high z-order base. They therefore form a separate z-stack above workbook drawing objects.
  const canvasChartCount = (() => {
    let count = 0;
    for (const { obj } of items) {
      if (isChartStoreDrawingId(obj.id)) count += 1;
    }
    return count;
  })();

  return (
    <div className="selection-pane" data-testid="selection-pane" tabIndex={0} ref={rootRef} onKeyDown={handleKeyDown}>
      {items.length === 0 ? (
        <div className="selection-pane__empty" data-testid="selection-pane-empty">
          No objects on this sheet.
        </div>
      ) : (
        <ul className="selection-pane__list" role="listbox" aria-label="Selection Pane objects">
          {items.map(({ obj, label }, index) => {
            const selected = obj.id === selectedId;
            const isCanvasChart = isChartStoreDrawingId(obj.id);
            const groupStart = isCanvasChart ? 0 : canvasChartCount;
            const groupSize = isCanvasChart ? canvasChartCount : Math.max(0, items.length - canvasChartCount);
            const groupIndex = index - groupStart;
            const canBringForward = groupIndex > 0;
            const canSendBackward = groupIndex >= 0 && groupIndex < groupSize - 1;
            return (
              <li
                key={obj.id}
                data-testid={`selection-pane-item-${obj.id}`}
                role="option"
                aria-selected={selected}
                className={selected ? "selection-pane__row selection-pane__row--selected" : "selection-pane__row"}
                onClick={() => {
                  // Match typical listbox behavior: clicking an option should focus the list
                  // so Arrow key navigation works immediately after click.
                  const root = rootRef.current;
                  if (root) {
                    try {
                      (root as any).focus?.({ preventScroll: true });
                    } catch {
                      root.focus?.();
                    }
                  }
                  app.selectDrawingById(obj.id);
                }}
              >
                <span className="selection-pane__icon" aria-hidden="true">
                  <DrawingKindIcon kind={obj.kind.type} />
                </span>
                <span className="selection-pane__label">{label}</span>
                {typeof app.bringSelectedDrawingForward === "function" ? (
                  <button
                    type="button"
                    className="selection-pane__action"
                    aria-label={`Bring forward ${label}`}
                    data-testid={`selection-pane-bring-forward-${obj.id}`}
                    disabled={actionsDisabled || !canBringForward}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      app.selectDrawingById(obj.id);
                      app.bringSelectedDrawingForward?.();
                    }}
                  >
                    <BringForwardIcon size={16} />
                  </button>
                ) : null}
                {typeof app.sendSelectedDrawingBackward === "function" ? (
                  <button
                    type="button"
                    className="selection-pane__action"
                    aria-label={`Send backward ${label}`}
                    data-testid={`selection-pane-send-backward-${obj.id}`}
                    disabled={actionsDisabled || !canSendBackward}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      app.selectDrawingById(obj.id);
                      app.sendSelectedDrawingBackward?.();
                    }}
                  >
                    <SendBackwardIcon size={16} />
                  </button>
                ) : null}
                <button
                  type="button"
                  className="selection-pane__action"
                  aria-label={`Delete ${label}`}
                  data-testid={`selection-pane-delete-${obj.id}`}
                  disabled={actionsDisabled}
                  onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    app.deleteDrawingById?.(obj.id);
                  }}
                >
                  <TrashIcon size={16} />
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}
