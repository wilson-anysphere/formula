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

  const focusFirstInput = (dialog) => {
    const input = dialog.querySelector?.("input");
    if (!input) return;
    const raf = globalThis.requestAnimationFrame;
    if (typeof raf === "function") {
      raf(() => {
        try {
          input.focus();
          input.select?.();
        } catch {
          // ignore
        }
      });
      return;
    }
    try {
      input.focus();
      input.select?.();
    } catch {
      // ignore
    }
  };

  const closeDialog = (dialog) => {
    try {
      if (typeof dialog.close === "function") dialog.close();
      else dialog.removeAttribute?.("open");
    } catch {
      // ignore
    }
  };

  const openDialog = (dialog) => {
    // Close siblings first so only one modal dialog is visible.
    for (const other of [findDialog, replaceDialog, goToDialog]) {
      if (other === dialog) continue;
      closeDialog(other);
    }
    try {
      if (typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute?.("open", "");
    } catch {
      // ignore
    }
    focusFirstInput(dialog);
  };

  const onKeyDown = (e) => {
    if (e.defaultPrevented) return;

    const key = String(e.key ?? "").toLowerCase();
    if (!key) return;

    // Replace
    // - Windows/Linux: Ctrl+H
    // - macOS: Cmd+Option+F (avoid Cmd+H which is the OS "Hide" shortcut)
    if ((e.ctrlKey && !e.metaKey && key === "h") || (e.metaKey && e.altKey && key === "f")) {
      e.preventDefault();
      openDialog(replaceDialog);
      return;
    }

    // Find: Ctrl+F / Cmd+F (but not Cmd+Option+F, which is Replace).
    if ((e.ctrlKey || e.metaKey) && !e.altKey && key === "f") {
      e.preventDefault();
      openDialog(findDialog);
      return;
    }

    // Go to: Ctrl+G / Cmd+G
    if ((e.ctrlKey || e.metaKey) && key === "g") {
      e.preventDefault();
      openDialog(goToDialog);
      return;
    }
  };

  window.addEventListener("keydown", onKeyDown);

  return { findDialog, replaceDialog, goToDialog };
}
