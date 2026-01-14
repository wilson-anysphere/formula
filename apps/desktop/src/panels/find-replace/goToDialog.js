import { formatA1Range, parseGoTo } from "../../../../../packages/search/index.js";
import { markKeybindingBarrier } from "../../keybindingBarrier.js";

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "className") node.className = v;
    else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
    else if (v != null) node.setAttribute(k, String(v));
  }
  for (const child of children) {
    node.append(child);
  }
  return node;
}

export function createGoToDialog({ workbook, getCurrentSheetName, setActiveCell, selectRange }) {
  const dialog = el("dialog", { className: "dialog goto-dialog" });
  dialog.dataset.testid = "goto-dialog";
  markKeybindingBarrier(dialog);
  const input = el("input", { type: "text", placeholder: "Reference (e.g. Sheet1!A1 or MyName)" });
  const error = el("div", { className: "error" });

  async function onGo() {
    error.textContent = "";
    try {
      const currentSheetName = getCurrentSheetName?.();
      const parsed = parseGoTo(input.value, { workbook, currentSheetName });
      if (parsed.type === "range") {
        const { range } = parsed;
        if (range.startRow === range.endRow && range.startCol === range.endCol) {
          setActiveCell?.({ sheetName: parsed.sheetName, row: range.startRow, col: range.startCol });
        } else {
          selectRange?.({ sheetName: parsed.sheetName, range });
        }
      }
      dialog.close();
    } catch (e) {
      error.textContent = e?.message ?? String(e);
    }
  }

  dialog.append(
    el("div", { className: "row" }, [input]),
    el("div", { className: "row" }, [
      el("button", { type: "button", onClick: onGo }, ["Go"]),
      el("button", { type: "button", onClick: () => dialog.close(), className: "goto-dialog__cancel" }, [
        "Cancel",
      ]),
    ]),
    error,
  );

  dialog.addEventListener("close", () => {
    input.value = "";
    error.textContent = "";
  });

  dialog.addEventListener("keydown", (e) => {
    if (e.key === "Enter") onGo();
  });

  // Useful for debugging in devtools.
  dialog.formatA1Range = formatA1Range;

  return dialog;
}
