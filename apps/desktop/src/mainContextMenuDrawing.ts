import type { SpreadsheetApp } from "./app/spreadsheetApp";
import { t } from "./i18n/index.js";
import type { ContextMenu, ContextMenuItem } from "./menus/contextMenu.js";
import type { DrawingObjectId } from "./drawings/types";

export type DrawingHitResult = { id: DrawingObjectId };

type DrawingContextMenuApp = Pick<
  SpreadsheetApp,
  | "hitTestDrawingAtClientPoint"
  | "getSelectedDrawingId"
  | "listDrawingsForSheet"
  | "isSelectedDrawingImage"
  | "selectDrawingById"
  | "cut"
  | "copy"
  | "deleteDrawingById"
  | "bringSelectedDrawingForward"
  | "sendSelectedDrawingBackward"
  | "focus"
>;

export function buildDrawingContextMenuItems(params: {
  app: DrawingContextMenuApp;
  isEditing: boolean;
}): ContextMenuItem[] {
  const { app, isEditing } = params;
  const selectedId = app.getSelectedDrawingId();
  const hasSelection = selectedId != null;
  const enabled = !isEditing && hasSelection;
  const clipboardEnabled = enabled && app.isSelectedDrawingImage();

  const { canBringForward, canSendBackward } = (() => {
    // Canvas charts use negative ids (stable hash namespace) and are not currently reorderable.
    if (!enabled || selectedId == null || selectedId < 0) return { canBringForward: false, canSendBackward: false };
    // `listDrawingsForSheet` returns topmost-first ordering.
    const drawings = app.listDrawingsForSheet();
    if (!Array.isArray(drawings) || drawings.length < 2) {
      return { canBringForward: false, canSendBackward: false };
    }
    const idx = drawings.findIndex((d) => d?.id === selectedId);
    if (idx < 0) return { canBringForward: false, canSendBackward: false };
    return { canBringForward: idx > 0, canSendBackward: idx < drawings.length - 1 };
  })();

  const cutLabelRaw = t("clipboard.cut");
  const cutLabel = cutLabelRaw === "clipboard.cut" ? "Cut" : cutLabelRaw;
  const copyLabelRaw = t("clipboard.copy");
  const copyLabel = copyLabelRaw === "clipboard.copy" ? "Copy" : copyLabelRaw;

  return [
    {
      type: "item",
      label: cutLabel,
      enabled: clipboardEnabled,
      onSelect: () => {
        // Ensure clipboard commands treat the grid as the active focus target (so
        // they don't early-return due to focus being in an input, and so Cut doesn't
        // restore focus back into the now-closed context menu).
        app.focus();
        app.cut();
      },
    },
    {
      type: "item",
      label: copyLabel,
      enabled: clipboardEnabled,
      onSelect: () => {
        app.focus();
        app.copy();
      },
    },
    {
      type: "item",
      label: "Delete",
      enabled,
      onSelect: () => {
        if (selectedId != null) {
          app.deleteDrawingById(selectedId);
        }
        app.focus();
      },
    },
    { type: "separator" },
    {
      type: "item",
      label: "Bring Forward",
      enabled: canBringForward,
      onSelect: () => {
        app.bringSelectedDrawingForward();
        app.focus();
      },
    },
    {
      type: "item",
      label: "Send Backward",
      enabled: canSendBackward,
      onSelect: () => {
        app.sendSelectedDrawingBackward();
        app.focus();
      },
    },
  ];
}

/**
 * Try to open a drawing-specific context menu at the provided client coordinates.
 *
 * Returns `true` when a drawing was hit and the drawing menu was opened.
 */
export function tryOpenDrawingContextMenuAtClientPoint(params: {
  app: DrawingContextMenuApp;
  contextMenu: ContextMenu;
  clientX: number;
  clientY: number;
  isEditing: boolean;
  onWillOpen?: () => void;
}): boolean {
  const { app, contextMenu, clientX, clientY, isEditing, onWillOpen } = params;

  const hit = app.hitTestDrawingAtClientPoint(clientX, clientY);
  if (!hit) return false;

  // Match Excel: right-click selects the object under the cursor without changing
  // the active cell selection.
  app.selectDrawingById(hit.id);

  onWillOpen?.();
  contextMenu.open({
    x: clientX,
    y: clientY,
    items: buildDrawingContextMenuItems({ app, isEditing }),
  });
  return true;
}
