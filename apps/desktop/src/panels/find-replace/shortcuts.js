import { createFindReplaceDialog } from "./findReplacePanel.js";
import { createGoToDialog } from "./goToDialog.js";

function isMod(e) {
  return e.ctrlKey || e.metaKey;
}

function showDialogWithFocus(dialog) {
  if (!dialog.open) dialog.showModal();
  requestAnimationFrame(() => {
    const input = dialog.querySelector("input, textarea");
    if (!input) return;
    input.focus();
    input.select?.();
  });
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
      showDialogWithFocus(findDialog);
    } else if (e.key.toLowerCase() === "h") {
      e.preventDefault();
      showDialogWithFocus(replaceDialog);
    } else if (e.key.toLowerCase() === "g") {
      e.preventDefault();
      showDialogWithFocus(goToDialog);
    }
  });

  return { findDialog, replaceDialog, goToDialog };
}
