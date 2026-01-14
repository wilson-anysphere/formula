import type { SpreadsheetApp } from "./app/spreadsheetApp";
import { t } from "./i18n/index.js";
import type { ContextMenu, ContextMenuItem } from "./menus/contextMenu.js";
import type { DrawingObjectId } from "./drawings/types";

export type DrawingHitResult = { id: DrawingObjectId };

type DrawingContextMenuApp = Pick<
  SpreadsheetApp,
  | "hitTestDrawingAtClientPoint"
  | "getSelectedDrawingId"
  | "selectDrawingById"
  | "cut"
  | "copy"
  | "deleteSelectedDrawing"
  | "bringSelectedDrawingForward"
  | "sendSelectedDrawingBackward"
  | "focus"
>;

export function buildDrawingContextMenuItems(params: {
  app: DrawingContextMenuApp;
  isEditing: boolean;
}): ContextMenuItem[] {
  const { app, isEditing } = params;
  const hasSelection = app.getSelectedDrawingId() != null;
  const enabled = !isEditing && hasSelection;

  const cutLabelRaw = t("clipboard.cut");
  const cutLabel = cutLabelRaw === "clipboard.cut" ? "Cut" : cutLabelRaw;
  const copyLabelRaw = t("clipboard.copy");
  const copyLabel = copyLabelRaw === "clipboard.copy" ? "Copy" : copyLabelRaw;

  return [
    {
      type: "item",
      label: cutLabel,
      enabled,
      onSelect: () => {
        app.cut();
        app.focus();
      },
    },
    {
      type: "item",
      label: copyLabel,
      enabled,
      onSelect: () => {
        app.copy();
        app.focus();
      },
    },
    {
      type: "item",
      label: "Delete",
      enabled,
      onSelect: () => {
        app.deleteSelectedDrawing();
        app.focus();
      },
    },
    { type: "separator" },
    {
      type: "item",
      label: "Bring Forward",
      enabled,
      onSelect: () => {
        app.bringSelectedDrawingForward();
        app.focus();
      },
    },
    {
      type: "item",
      label: "Send Backward",
      enabled,
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
