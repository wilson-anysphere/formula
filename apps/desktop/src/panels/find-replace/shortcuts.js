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
    // Replace: Cmd+Option+F on macOS (Cmd+H is reserved for Hide).
    const key = e.key.toLowerCase();
    if (e.metaKey && e.altKey && (key === "f" || e.code === "KeyF")) {
      e.preventDefault();
      showDialogWithFocus(replaceDialog);
      return;
    }

    if (!isMod(e)) return;

    if (key === "f") {
      e.preventDefault();
      showDialogWithFocus(findDialog);
    } else if (key === "h" && e.ctrlKey) {
      e.preventDefault();
      showDialogWithFocus(replaceDialog);
    } else if (key === "g") {
      e.preventDefault();
      showDialogWithFocus(goToDialog);
    }
  });

  return { findDialog, replaceDialog, goToDialog };
}
