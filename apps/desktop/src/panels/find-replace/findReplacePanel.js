import { formatA1Address } from "../../../../../packages/search/index.js";

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
  const dialog = el("dialog", { className: "find-replace-dialog" });

  const findInput = el("input", { type: "text", placeholder: "Find what…" });
  const replaceInput = el("input", { type: "text", placeholder: "Replace with…" });
  replaceInput.style.display = mode === "replace" ? "" : "none";

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
      const item = el(
        "li",
        {
          onClick: () =>
            controller.setActiveCell?.({ sheetName: r.sheetName, row: r.row, col: r.col }),
        },
        [`${r.sheetName}!${formatA1Address({ row: r.row, col: r.col })} — ${r.text}`],
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
      el("label", { style: "margin-left: 8px" }, ["Look in: "]),
      lookInSelect,
    ]),
    el("div", { className: "row" }, [
      el("label", {}, [matchCase, " Match case"]),
      el("label", { style: "margin-left: 12px" }, [wholeCell, " Match entire cell"]),
    ]),
    el("div", { className: "row" }, [
      el("button", { onClick: onFindNext, type: "button" }, ["Find Next"]),
      el("button", { onClick: onFindAll, type: "button" }, ["Find All"]),
      el(
        "button",
        { onClick: onReplaceNext, type: "button", style: mode === "replace" ? "" : "display:none" },
        ["Replace"],
      ),
      el(
        "button",
        { onClick: onReplaceAll, type: "button", style: mode === "replace" ? "" : "display:none" },
        ["Replace All"],
      ),
      el(
        "button",
        { onClick: () => dialog.close(), type: "button", style: "margin-left: auto" },
        ["Close"],
      ),
    ]),
    resultsList,
  );

  return dialog;
}
