import { createFindReplaceDialog } from "./findReplacePanel.js";
import { createGoToDialog } from "./goToDialog.js";

export function registerFindReplaceShortcuts({
  controller,
  workbook,
  getCurrentSheetName,
  setActiveCell,
  selectRange,
  mount = document.body,
}) {
  const findDialog = createFindReplaceDialog(controller, { mode: "find" });
  const replaceDialog = createFindReplaceDialog(controller, { mode: "replace" });
  const goToDialog = createGoToDialog({ workbook, getCurrentSheetName, setActiveCell, selectRange });

  mount.append(findDialog, replaceDialog, goToDialog);

  return { findDialog, replaceDialog, goToDialog };
}

