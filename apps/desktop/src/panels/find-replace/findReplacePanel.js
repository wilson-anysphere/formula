import { formatA1Address } from "../../../../../packages/search/index.js";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.ts";
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

export function createFindReplaceDialog(controller, { mode = "find" } = {}) {
  const dialog = el("dialog", { className: "dialog find-replace-dialog" });
  dialog.dataset.testid = mode === "replace" ? "replace-dialog" : "find-dialog";
  markKeybindingBarrier(dialog);

  const findInput = el("input", { type: "text", placeholder: "Find what…" });
  const replaceInput = el("input", { type: "text", placeholder: "Replace with…" });
  replaceInput.hidden = mode !== "replace";

  const scopeSelect = el(
    "select",
    {},
    ["selection", "sheet", "workbook"].map((v) => el("option", { value: v }, [v])),
  );
  scopeSelect.value = controller.scope;

  const lookInSelect = el(
    "select",
    {},
    ["values", "formulas"].map((v) => el("option", { value: v }, [v])),
  );
  lookInSelect.value = controller.lookIn;

  const matchCase = el("input", { type: "checkbox" });
  const wholeCell = el("input", { type: "checkbox" });

  const resultsList = el("ul", { className: "find-results" });

  function setControllerFromUI() {
    controller.query = findInput.value;
    controller.replacement = replaceInput.value;
    controller.scope = scopeSelect.value;
    controller.lookIn = lookInSelect.value;
    controller.matchCase = matchCase.checked;
    controller.matchEntireCell = wholeCell.checked;
  }

  async function onFindNext() {
    setControllerFromUI();
    await controller.findNext();
  }

  async function onFindAll() {
    setControllerFromUI();
    const results = await controller.findAll();
    resultsList.replaceChildren();
    for (const r of results) {
      const sheetToken = formatSheetNameForA1(r.sheetName);
      const sheetPrefix = sheetToken ? `${sheetToken}!` : "";
      const item = el(
        "li",
        {
          onClick: () =>
            controller.setActiveCell?.({ sheetName: r.sheetName, row: r.row, col: r.col }),
        },
        [`${sheetPrefix}${formatA1Address({ row: r.row, col: r.col })} — ${r.text}`],
      );
      resultsList.append(item);
    }
  }

  async function onReplaceNext() {
    setControllerFromUI();
    await controller.replaceNext();
  }

  async function onReplaceAll() {
    setControllerFromUI();
    await controller.replaceAll();
    await onFindAll();
  }

  dialog.append(
    el("div", { className: "row" }, [el("label", {}, ["Find: "]), findInput]),
    el("div", { className: "row" }, [el("label", {}, ["Replace: "]), replaceInput]),
    el("div", { className: "row" }, [
      el("label", {}, ["Within: "]),
      scopeSelect,
      el("label", { className: "find-replace-dialog__look-in-label" }, ["Look in: "]),
      lookInSelect,
    ]),
    el("div", { className: "row" }, [
      el("label", {}, [matchCase, " Match case"]),
      el("label", { className: "find-replace-dialog__whole-cell-label" }, [wholeCell, " Match entire cell"]),
    ]),
    (() => {
      const row = el("div", { className: "row" });
      row.append(el("button", { onClick: onFindNext, type: "button" }, ["Find Next"]));
      row.append(el("button", { onClick: onFindAll, type: "button" }, ["Find All"]));

      const replaceBtn = el("button", { onClick: onReplaceNext, type: "button" }, ["Replace"]);
      replaceBtn.hidden = mode !== "replace";
      row.append(replaceBtn);

      const replaceAllBtn = el("button", { onClick: onReplaceAll, type: "button" }, ["Replace All"]);
      replaceAllBtn.hidden = mode !== "replace";
      row.append(replaceAllBtn);

      row.append(el("button", { onClick: () => dialog.close(), type: "button", className: "find-replace-dialog__close" }, ["Close"]));
      return row;
    })(),
    resultsList,
  );

  return dialog;
}
