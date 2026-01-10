import { createFindReplaceDialog } from "./findReplacePanel.js";
import { createGoToDialog } from "./goToDialog.js";

function isMod(e) {
  return e.ctrlKey || e.metaKey;
}

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

  window.addEventListener("keydown", (e) => {
    if (!isMod(e)) return;

    if (e.key.toLowerCase() === "f") {
      e.preventDefault();
      findDialog.showModal();
    } else if (e.key.toLowerCase() === "h") {
      e.preventDefault();
      replaceDialog.showModal();
    } else if (e.key.toLowerCase() === "g") {
      e.preventDefault();
      goToDialog.showModal();
    }
  });

  return { findDialog, replaceDialog, goToDialog };
}
