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

  const closeAllExcept = (keep) => {
    for (const dialog of [findDialog, replaceDialog, goToDialog]) {
      if (dialog === keep) continue;
      if (typeof dialog.close === "function" && dialog.open) {
        dialog.close();
      }
    }
  };

  const showDialog = (dialog) => {
    if (!dialog) return;
    closeAllExcept(dialog);
    if (!dialog.open && typeof dialog.showModal === "function") {
      dialog.showModal();
    }
  };

  // Lightweight global shortcuts (Excel-esque defaults):
  // - Find: Cmd+F / Ctrl+F
  // - Replace: Cmd+Option+F / Ctrl+H
  //
  // macOS: avoid Cmd+H which is the system "Hide" shortcut.
  window.addEventListener("keydown", (e) => {
    if (!e || e.defaultPrevented) return;
    const key = String(e.key ?? "").toLowerCase();
    if (!key) return;

    // Cmd+H is reserved for "Hide" on macOS; do not bind Replace to it.
    if (e.metaKey && !e.ctrlKey && !e.altKey && key === "h") {
      return;
    }

    // Replace (mac): Cmd+Option+F
    if (e.metaKey && e.altKey && !e.ctrlKey && key === "f") {
      e.preventDefault();
      showDialog(replaceDialog);
      return;
    }

    // Replace (win/linux): Ctrl+H
    if (e.ctrlKey && !e.metaKey && key === "h") {
      e.preventDefault();
      showDialog(replaceDialog);
      return;
    }

    // Find: Cmd+F / Ctrl+F
    if ((e.metaKey && !e.ctrlKey && !e.altKey && key === "f") || (e.ctrlKey && !e.metaKey && key === "f")) {
      e.preventDefault();
      showDialog(findDialog);
      return;
    }
  });

  return { findDialog, replaceDialog, goToDialog };
}
