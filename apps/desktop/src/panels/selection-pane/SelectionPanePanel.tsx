import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { DrawingObject, DrawingObjectId } from "../../drawings/types";

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

export function SelectionPanePanel({ app }: { app: SelectionPaneApp }) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [drawings, setDrawings] = useState<DrawingObject[]>(() => app.listDrawingsForSheet());
  const [selectedId, setSelectedId] = useState<DrawingObjectId | null>(() => app.getSelectedDrawingId());

  useEffect(() => {
    const onDrawings = () => {
      setDrawings(app.listDrawingsForSheet());
    };
    return app.subscribeDrawings(onDrawings);
  }, [app]);

  useEffect(() => {
    const onSelection = (id: DrawingObjectId | null) => {
      setSelectedId(id);
    };
    return app.subscribeDrawingSelection(onSelection);
  }, [app]);

  const items = useMemo(() => {
    // Excel-style auto-naming: `Picture 1`, `Shape 1`, `Chart 1` etc.
    // Keep this stable for a given list ordering (topmost-first from `listDrawingsForSheet`).
    let picture = 0;
    let shape = 0;
    let chart = 0;
    let unknown = 0;
    return drawings.map((obj) => {
      const kind = obj.kind.type;
      const explicitLabel = (obj.kind as { label?: string }).label?.trim();
      if (explicitLabel) {
        return { obj, label: explicitLabel };
      }
      switch (kind) {
        case "image":
          picture += 1;
          return { obj, label: `Picture ${picture}` };
        case "shape":
          shape += 1;
          return { obj, label: `Shape ${shape}` };
        case "chart":
          chart += 1;
          return { obj, label: `Chart ${chart}` };
        default:
          unknown += 1;
          return { obj, label: `${kind} ${unknown} (id=${obj.id})` };
      }
    });
  }, [drawings]);

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

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      if (items.length === 0) return;
      // Only handle keys when the Selection Pane root itself is focused.
      // (If a per-row action button is focused, allow normal Tab navigation and avoid
      // hijacking keyboard interactions.)
      if (e.target !== e.currentTarget) return;
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
          selectIndex(currentIndex < 0 ? 0 : Math.min(currentIndex + 1, items.length - 1));
          return;
        case "ArrowUp":
          selectIndex(currentIndex < 0 ? items.length - 1 : Math.max(currentIndex - 1, 0));
          return;
        case "Home":
          selectIndex(0);
          return;
        case "End":
          selectIndex(items.length - 1);
          return;
        case "Delete":
        case "Backspace": {
          if (selectedId == null) return;
          if (typeof app.deleteDrawingById !== "function") return;
          e.preventDefault();
          e.stopPropagation();
          app.deleteDrawingById(selectedId);
          return;
        }
        default:
          return;
      }
    },
    [app, items, scrollItemIntoView, selectedId],
  );

  return (
    <div className="selection-pane" data-testid="selection-pane" tabIndex={0} ref={rootRef} onKeyDown={handleKeyDown}>
      {items.length === 0 ? (
        <div className="selection-pane__empty" data-testid="selection-pane-empty">
          No objects on this sheet.
        </div>
      ) : (
        <ul className="selection-pane__list" role="listbox" aria-label="Selection Pane objects">
          {items.map(({ obj, label }) => {
            const selected = obj.id === selectedId;
            return (
              <li
                key={obj.id}
                data-testid={`selection-pane-item-${obj.id}`}
                role="option"
                aria-selected={selected}
                className={selected ? "selection-pane__row selection-pane__row--selected" : "selection-pane__row"}
                onClick={() => app.selectDrawingById(obj.id)}
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
