import { createFindReplaceDialog } from "./findReplacePanel.js";
import { createGoToDialog } from "./goToDialog.js";

function isTextInputLike(target) {
  const el = target;
  if (!el || typeof el !== "object") return false;
  const tag = el.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || el.isContentEditable;
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
    const focusInput = () => {
      const input = dialog.querySelector("input, textarea");
      if (!input) return;
      input.focus();
      if (typeof input.select === "function") input.select();
    };
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(focusInput);
    } else {
      setTimeout(focusInput, 0);
    }
  };

  // Lightweight global shortcuts (Excel-esque defaults):
  // - Find: Cmd+F / Ctrl+F
  // - Replace: Cmd+Option+F / Ctrl+H
  // - Go To: Cmd+G / Ctrl+G
  //
  // macOS: avoid Cmd+H which is the system "Hide" shortcut.
  window.addEventListener("keydown", (e) => {
    if (!e || e.defaultPrevented) return;
    if (isTextInputLike(e.target)) return;

    const key = String(e.key ?? "").toLowerCase();
    if (!key) return;

    // Cmd+H is reserved for "Hide" on macOS; do not bind Replace to it.
    if (e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey && key === "h") {
      return;
    }

    // Replace (mac): Cmd+Option+F
    if (e.metaKey && e.altKey && !e.ctrlKey && !e.shiftKey && key === "f") {
      e.preventDefault();
      showDialog(replaceDialog);
      return;
    }

    // Replace (win/linux): Ctrl+H
    if (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && key === "h") {
      e.preventDefault();
      showDialog(replaceDialog);
      return;
    }

    // Find: Cmd+F / Ctrl+F
    if (
      (e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey && key === "f") ||
      (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && key === "f")
    ) {
      e.preventDefault();
      showDialog(findDialog);
      return;
    }

    // Go To: Cmd+G / Ctrl+G
    if (
      (e.metaKey && !e.ctrlKey && !e.altKey && !e.shiftKey && key === "g") ||
      (e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey && key === "g")
    ) {
      e.preventDefault();
      showDialog(goToDialog);
    }
  });

  return { findDialog, replaceDialog, goToDialog };
}
