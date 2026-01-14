import { applyFormatCells } from "./formatCellsDialog.js";
import { getEffectiveCellStyle } from "./getEffectiveCellStyle.js";
import { getStyleFillFgColor, getStyleFontSizePt, getStyleNumberFormat, getStyleWrapText } from "./styleFieldAccess.js";
import { showToast } from "../extensions/ui.js";
import { markKeybindingBarrier } from "../keybindingBarrier.js";
import { DEFAULT_GRID_LIMITS } from "../selection/selection.js";
import type { GridLimits, Range } from "../selection/types";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT, evaluateFormattingSelectionSize, normalizeSelectionRange } from "./selectionSizeGuard.js";
import { resolveCssVar } from "../theme/cssVars.js";

export type FormatCellsDialogHost = {
  isEditing: () => boolean;
  /**
   * Optional read-only indicator (used in collab viewer/commenter sessions).
   *
   * In read-only roles we only allow applying formatting when the selection is an entire
   * row/column/sheet band (i.e. "formatting defaults"). These mutations are intended to be
   * local-only; collab binders prevent persisting them into the shared document state.
   */
  isReadOnly?: () => boolean;
  getDocument: () => any;
  getSheetId: () => string;
  getActiveCell: () => { row: number; col: number };
  getSelectionRanges: () => Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>;
  /**
   * Optional: the current grid limits (used for band detection + expanding full row/col selections
   * to canonical Excel bounds so layered formatting fast paths apply).
   */
  getGridLimits?: () => GridLimits;
  focusGrid: () => void;
};

const NUMBER_FORMAT_CODE_BY_PRESET: Record<string, string> = {
  currency: "$#,##0.00",
  percent: "0%",
  date: "m/d/yyyy",
};

const FILL_SWATCH_CSS_VAR_BY_ID: Record<string, string> = {
  series1: "--chart-series-1",
  series2: "--chart-series-2",
  series3: "--chart-series-3",
  series4: "--chart-series-4",
};

function resolveCssVarValue(varName: string): string {
  return resolveCssVar(varName, { fallback: "" });
}

function cssHexToArgb(hex: string): string | null {
  const match = hex.trim().match(/^#([0-9a-f]{6})$/i);
  if (!match) return null;
  return `#FF${match[1]!.toUpperCase()}`;
}

function resolveFillArgbFromSwatchId(id: string): string | null {
  const cssVar = FILL_SWATCH_CSS_VAR_BY_ID[id];
  if (!cssVar) return null;
  const value = resolveCssVarValue(cssVar);
  return cssHexToArgb(value);
}

function normalizeArgb(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  return trimmed.toUpperCase();
}

function showDialogModal(dialog: HTMLDialogElement): void {
  // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
  if (typeof dialog.showModal === "function") {
    try {
      // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
      dialog.showModal();
      return;
    } catch {
      // Fall through to non-modal open attribute.
    }
  }
  // jsdom fallback: `open` attribute is enough for our tests.
  dialog.setAttribute("open", "true");
}

export function openFormatCellsDialog(host: FormatCellsDialogHost): void {
  if (host.isEditing()) return;

  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) {
    if (openModal.classList.contains("format-cells-dialog")) return;
    return;
  }

  const doc = host.getDocument();
  const sheetId = host.getSheetId();
  const active = host.getActiveCell();
  const activeStyle = getEffectiveCellStyle(doc, sheetId, { row: active.row, col: active.col });
  const isReadOnlyAtOpen = host.isReadOnly?.() === true;

  if (isReadOnlyAtOpen) {
    const selectionRanges = host.getSelectionRanges();
    const rawRanges =
      selectionRanges.length > 0
        ? selectionRanges
        : [{ startRow: active.row, startCol: active.col, endRow: active.row, endCol: active.col }];
    const limits: GridLimits = host.getGridLimits?.() ?? DEFAULT_GRID_LIMITS;
    const decision = evaluateFormattingSelectionSize(rawRanges as Range[], limits, {
      maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT,
    });
    if (!decision.allRangesBand) {
      try {
        showToast("Read-only: select an entire row, column, or sheet to change formatting defaults.", "warning");
      } catch {
        // ignore (e.g. toast root missing in tests)
      }
      return;
    }
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog format-cells-dialog";
  dialog.dataset.testid = "format-cells-dialog";
  markKeybindingBarrier(dialog);

  const header = document.createElement("div");
  header.className = "format-cells-dialog__header";
  const title = document.createElement("h3");
  title.className = "format-cells-dialog__title";
  title.textContent = "Format Cells";
  header.appendChild(title);
  dialog.appendChild(header);

  const content = document.createElement("div");
  content.className = "format-cells-dialog__content";
  dialog.appendChild(content);

  // --- Number ----------------------------------------------------------------

  const numberSection = document.createElement("section");
  numberSection.className = "format-cells-dialog__section";
  const numberTitle = document.createElement("h4");
  numberTitle.className = "format-cells-dialog__section-title";
  numberTitle.textContent = "Number";
  numberSection.appendChild(numberTitle);

  const numberRow = document.createElement("div");
  numberRow.className = "format-cells-dialog__row";
  const numberLabel = document.createElement("div");
  numberLabel.className = "format-cells-dialog__label";
  numberLabel.textContent = "Format";
  const numberSelectWrap = document.createElement("div");
  numberSelectWrap.className = "format-cells-dialog__control";
  const numberSelect = document.createElement("select");
  numberSelect.className = "format-cells-dialog__select";
  numberSelect.dataset.testid = "format-cells-number";
  const numberOptions: Array<{ value: string; label: string }> = [
    { value: "", label: "General" },
    { value: "currency", label: "Currency" },
    { value: "percent", label: "Percent" },
    { value: "date", label: "Date" },
    { value: "__custom__", label: "Custom" },
  ];
  for (const opt of numberOptions) {
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.label;
    numberSelect.appendChild(o);
  }
  numberSelectWrap.appendChild(numberSelect);
  numberRow.appendChild(numberLabel);
  numberRow.appendChild(numberSelectWrap);
  numberSection.appendChild(numberRow);

  const numberCustomRow = document.createElement("div");
  numberCustomRow.className = "format-cells-dialog__row";
  const numberCustomLabel = document.createElement("div");
  numberCustomLabel.className = "format-cells-dialog__label";
  numberCustomLabel.textContent = "Code";
  const numberCustomWrap = document.createElement("div");
  numberCustomWrap.className = "format-cells-dialog__control";
  const numberCustomInput = document.createElement("input");
  numberCustomInput.className = "format-cells-dialog__input";
  numberCustomInput.type = "text";
  numberCustomInput.placeholder = "General";
  numberCustomInput.dataset.testid = "format-cells-number-custom";
  numberCustomWrap.appendChild(numberCustomInput);
  numberCustomRow.appendChild(numberCustomLabel);
  numberCustomRow.appendChild(numberCustomWrap);
  numberSection.appendChild(numberCustomRow);

  const syncNumberCustomVisibility = () => {
    const isCustom = numberSelect.value === "__custom__";
    numberCustomRow.style.display = isCustom ? "" : "none";
    // When the custom row is hidden, ensure we don't leave an extra bottom margin on the only
    // visible row in the section (since `:last-child` won't match due to the hidden node).
    numberRow.style.marginBottom = isCustom ? "" : "0";
    if (isCustom) {
      // When switching to Custom, move focus to the code field so users can type immediately.
      try {
        numberCustomInput.focus();
      } catch {
        // ignore (e.g. jsdom)
      }
    }
  };
  numberSelect.addEventListener("change", syncNumberCustomVisibility);
  content.appendChild(numberSection);

  // --- Font ------------------------------------------------------------------

  const fontSection = document.createElement("section");
  fontSection.className = "format-cells-dialog__section";
  const fontTitle = document.createElement("h4");
  fontTitle.className = "format-cells-dialog__section-title";
  fontTitle.textContent = "Font";
  fontSection.appendChild(fontTitle);

  const fontToggleRow = document.createElement("div");
  fontToggleRow.className = "format-cells-dialog__row";
  const fontToggleLabel = document.createElement("div");
  fontToggleLabel.className = "format-cells-dialog__label";
  fontToggleLabel.textContent = "Style";
  const fontToggleWrap = document.createElement("div");
  fontToggleWrap.className = "format-cells-dialog__control";
  const toggles = document.createElement("div");
  toggles.className = "format-cells-dialog__toggles";

  const makeToggle = (label: string, testid: string) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "format-cells-dialog__toggle";
    btn.textContent = label;
    btn.dataset.testid = testid;
    btn.setAttribute("aria-pressed", "false");
    btn.addEventListener("click", () => {
      const pressed = btn.getAttribute("aria-pressed") === "true";
      btn.setAttribute("aria-pressed", pressed ? "false" : "true");
    });
    return btn;
  };

  const boldBtn = makeToggle("B", "format-cells-bold");
  const italicBtn = makeToggle("I", "format-cells-italic");
  const underlineBtn = makeToggle("U", "format-cells-underline");
  toggles.appendChild(boldBtn);
  toggles.appendChild(italicBtn);
  toggles.appendChild(underlineBtn);
  fontToggleWrap.appendChild(toggles);
  fontToggleRow.appendChild(fontToggleLabel);
  fontToggleRow.appendChild(fontToggleWrap);
  fontSection.appendChild(fontToggleRow);

  const fontSizeRow = document.createElement("div");
  fontSizeRow.className = "format-cells-dialog__row";
  const fontSizeLabel = document.createElement("div");
  fontSizeLabel.className = "format-cells-dialog__label";
  fontSizeLabel.textContent = "Size (pt)";
  const fontSizeWrap = document.createElement("div");
  fontSizeWrap.className = "format-cells-dialog__control";
  const fontSizeInput = document.createElement("input");
  fontSizeInput.className = "format-cells-dialog__input";
  fontSizeInput.type = "number";
  fontSizeInput.min = "1";
  fontSizeInput.step = "1";
  fontSizeInput.inputMode = "numeric";
  fontSizeInput.dataset.testid = "format-cells-font-size";
  fontSizeWrap.appendChild(fontSizeInput);
  fontSizeRow.appendChild(fontSizeLabel);
  fontSizeRow.appendChild(fontSizeWrap);
  fontSection.appendChild(fontSizeRow);

  content.appendChild(fontSection);

  // --- Fill ------------------------------------------------------------------

  const fillSection = document.createElement("section");
  fillSection.className = "format-cells-dialog__section";
  const fillTitle = document.createElement("h4");
  fillTitle.className = "format-cells-dialog__section-title";
  fillTitle.textContent = "Fill";
  fillSection.appendChild(fillTitle);

  const fillRow = document.createElement("div");
  fillRow.className = "format-cells-dialog__row";
  const fillLabel = document.createElement("div");
  fillLabel.className = "format-cells-dialog__label";
  fillLabel.textContent = "Color";
  const fillWrap = document.createElement("div");
  fillWrap.className = "format-cells-dialog__control";
  const swatches = document.createElement("div");
  swatches.className = "format-cells-dialog__swatches";

  let selectedFill = "none";

  const swatchDefs: Array<{ id: string; label: string }> = [
    { id: "none", label: "No fill" },
    { id: "custom", label: "Preserve existing fill" },
    { id: "series1", label: "Fill color 1" },
    { id: "series2", label: "Fill color 2" },
    { id: "series3", label: "Fill color 3" },
    { id: "series4", label: "Fill color 4" },
  ];

  const swatchButtons: HTMLButtonElement[] = [];

  const syncSwatchSelection = () => {
    for (const btn of swatchButtons) {
      btn.dataset.selected = btn.dataset.color === selectedFill ? "true" : "false";
    }
  };

  for (const swatch of swatchDefs) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "format-cells-dialog__swatch";
    btn.dataset.color = swatch.id;
    btn.setAttribute("aria-label", swatch.label);
    btn.addEventListener("click", () => {
      selectedFill = swatch.id;
      syncSwatchSelection();
    });
    swatchButtons.push(btn);
    swatches.appendChild(btn);
  }
  syncSwatchSelection();

  fillWrap.appendChild(swatches);
  fillRow.appendChild(fillLabel);
  fillRow.appendChild(fillWrap);
  fillSection.appendChild(fillRow);
  content.appendChild(fillSection);

  // --- Alignment -------------------------------------------------------------

  const alignSection = document.createElement("section");
  alignSection.className = "format-cells-dialog__section";
  const alignTitle = document.createElement("h4");
  alignTitle.className = "format-cells-dialog__section-title";
  alignTitle.textContent = "Alignment";
  alignSection.appendChild(alignTitle);

  const horizRow = document.createElement("div");
  horizRow.className = "format-cells-dialog__row";
  const horizLabel = document.createElement("div");
  horizLabel.className = "format-cells-dialog__label";
  horizLabel.textContent = "Horizontal";
  const horizWrap = document.createElement("div");
  horizWrap.className = "format-cells-dialog__control";
  const horizSelect = document.createElement("select");
  horizSelect.className = "format-cells-dialog__select";
  horizSelect.dataset.testid = "format-cells-horizontal";
  const horizOptions: Array<{ value: string; label: string }> = [
    { value: "", label: "General" },
    { value: "left", label: "Left" },
    { value: "center", label: "Center" },
    { value: "right", label: "Right" },
    { value: "__custom__", label: "Custom" },
  ];
  for (const opt of horizOptions) {
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.label;
    horizSelect.appendChild(o);
  }
  horizWrap.appendChild(horizSelect);
  horizRow.appendChild(horizLabel);
  horizRow.appendChild(horizWrap);
  alignSection.appendChild(horizRow);

  const wrapRow = document.createElement("div");
  wrapRow.className = "format-cells-dialog__row";
  const wrapLabel = document.createElement("div");
  wrapLabel.className = "format-cells-dialog__label";
  wrapLabel.textContent = "Text";
  const wrapWrap = document.createElement("div");
  wrapWrap.className = "format-cells-dialog__control";
  const wrapToggleLabel = document.createElement("label");
  wrapToggleLabel.className = "format-cells-dialog__checkbox";
  const wrapInput = document.createElement("input");
  wrapInput.type = "checkbox";
  wrapInput.dataset.testid = "format-cells-wrap";
  wrapToggleLabel.appendChild(wrapInput);
  wrapToggleLabel.appendChild(document.createTextNode("Wrap Text"));
  wrapWrap.appendChild(wrapToggleLabel);
  wrapRow.appendChild(wrapLabel);
  wrapRow.appendChild(wrapWrap);
  alignSection.appendChild(wrapRow);

  content.appendChild(alignSection);

  // --- Buttons ---------------------------------------------------------------

  const actions = document.createElement("div");
  actions.className = "dialog__controls format-cells-dialog__actions";
  dialog.appendChild(actions);

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = "Cancel";
  cancelBtn.dataset.testid = "format-cells-cancel";

  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.textContent = "OK";
  okBtn.dataset.testid = "format-cells-ok";
  okBtn.dataset.primary = "true";

  const applyBtn = document.createElement("button");
  applyBtn.type = "button";
  applyBtn.textContent = "Apply";
  applyBtn.dataset.testid = "format-cells-apply";

  actions.appendChild(cancelBtn);
  actions.appendChild(applyBtn);
  actions.appendChild(okBtn);

  // --- Initialize UI from active style --------------------------------------

  const activeNumberFormat = getStyleNumberFormat(activeStyle);
  numberCustomInput.value = activeNumberFormat ?? "";
  const initialPreset =
    activeNumberFormat == null
      ? ""
      : Object.entries(NUMBER_FORMAT_CODE_BY_PRESET).find(([, code]) => code === activeNumberFormat)?.[0] ?? "__custom__";
  numberSelect.value = initialPreset;
  syncNumberCustomVisibility();

  boldBtn.setAttribute("aria-pressed", Boolean(activeStyle?.font?.bold) ? "true" : "false");
  italicBtn.setAttribute("aria-pressed", Boolean(activeStyle?.font?.italic) ? "true" : "false");
  underlineBtn.setAttribute("aria-pressed", Boolean(activeStyle?.font?.underline) ? "true" : "false");

  const activeFontSize = getStyleFontSizePt(activeStyle);
  fontSizeInput.value = typeof activeFontSize === "number" ? String(activeFontSize) : "";

  const activeFill = normalizeArgb(getStyleFillFgColor(activeStyle));
  if (!activeFill) {
    selectedFill = "none";
  } else {
    selectedFill = "custom";
    for (const swatch of swatchDefs) {
      if (swatch.id === "none" || swatch.id === "custom") continue;
      const argb = normalizeArgb(resolveFillArgbFromSwatchId(swatch.id));
      if (argb && argb === activeFill) {
        selectedFill = swatch.id;
        break;
      }
    }
  }
  syncSwatchSelection();

  const activeHorizontal = typeof activeStyle?.alignment?.horizontal === "string" ? String(activeStyle.alignment.horizontal) : "";
  horizSelect.value =
    activeHorizontal === "left" || activeHorizontal === "center" || activeHorizontal === "right"
      ? activeHorizontal
      : activeHorizontal
        ? "__custom__"
        : "";
  wrapInput.checked = getStyleWrapText(activeStyle);

  function computeChanges(currentStyle: any): Record<string, any> | null {
    /** @type {Record<string, any>} */
    const changes: Record<string, any> = {};

    // Number
    const preset = numberSelect.value;
    const desiredNumberFormat = (() => {
      if (preset === "__custom__") {
        const raw = numberCustomInput.value;
        const trimmed = raw.trim();
        // Treat empty/"General" (Excel semantics) as clearing the number format.
        if (!trimmed || trimmed.toLowerCase() === "general") return null;
        // Preserve exact user input; avoid trimming away intentional whitespace in format codes.
        return raw;
      }
      return preset ? NUMBER_FORMAT_CODE_BY_PRESET[preset] ?? null : null;
    })();
    const currentNumberFormat = getStyleNumberFormat(currentStyle);
    if ((currentNumberFormat ?? null) !== (desiredNumberFormat ?? null)) {
      changes.numberFormat = desiredNumberFormat;
    }

    // Font
    const fontPatch: Record<string, any> = {};
    const desiredBold = boldBtn.getAttribute("aria-pressed") === "true";
    const desiredItalic = italicBtn.getAttribute("aria-pressed") === "true";
    const desiredUnderline = underlineBtn.getAttribute("aria-pressed") === "true";
    if (Boolean(currentStyle?.font?.bold) !== desiredBold) fontPatch.bold = desiredBold;
    if (Boolean(currentStyle?.font?.italic) !== desiredItalic) fontPatch.italic = desiredItalic;
    if (Boolean(currentStyle?.font?.underline) !== desiredUnderline) fontPatch.underline = desiredUnderline;

    const parsedSize = Number(fontSizeInput.value);
    const desiredSize = Number.isFinite(parsedSize) && parsedSize > 0 ? parsedSize : null;
    const currentSize = getStyleFontSizePt(currentStyle);
    if ((currentSize ?? null) !== (desiredSize ?? null)) fontPatch.size = desiredSize;

    if (Object.keys(fontPatch).length > 0) changes.font = fontPatch;

    // Fill
    if (selectedFill !== "custom") {
      const desiredFillArgb = selectedFill === "none" ? null : resolveFillArgbFromSwatchId(selectedFill);
      const currentFillArgb = normalizeArgb(getStyleFillFgColor(currentStyle));
      if ((currentFillArgb ?? null) !== (normalizeArgb(desiredFillArgb) ?? null)) {
        changes.fill = desiredFillArgb ? { pattern: "solid", fgColor: desiredFillArgb } : null;
      }
    }

    // Alignment
    const alignmentPatch: Record<string, any> = {};
    if (horizSelect.value !== "__custom__") {
      const desiredHorizontal = horizSelect.value ? horizSelect.value : null;
      const currentHorizontal =
        typeof currentStyle?.alignment?.horizontal === "string" ? currentStyle.alignment.horizontal : null;
      if ((currentHorizontal ?? null) !== (desiredHorizontal ?? null)) {
        alignmentPatch.horizontal = desiredHorizontal;
      }
    }

    const desiredWrap = wrapInput.checked;
    const currentWrap = getStyleWrapText(currentStyle);
    if (currentWrap !== desiredWrap) alignmentPatch.wrapText = desiredWrap;
    if (Object.keys(alignmentPatch).length > 0) changes.alignment = alignmentPatch;

    return Object.keys(changes).length > 0 ? changes : null;
  }

  function applyFromUi(): void {
    const sheetIdNow = host.getSheetId();
    const activeNow = host.getActiveCell();
    const styleNow = getEffectiveCellStyle(doc, sheetIdNow, { row: activeNow.row, col: activeNow.col });
    const changes = computeChanges(styleNow);
    if (!changes) return;

    const isReadOnly = host.isReadOnly?.() === true;
    const selectionRanges = host.getSelectionRanges();
    const rawRanges =
      selectionRanges.length > 0
        ? selectionRanges
        : [{ startRow: activeNow.row, startCol: activeNow.col, endRow: activeNow.row, endCol: activeNow.col }];

    const limits: GridLimits = host.getGridLimits?.() ?? DEFAULT_GRID_LIMITS;
    const decision = evaluateFormattingSelectionSize(rawRanges as Range[], limits, {
      maxCells: DEFAULT_FORMATTING_APPLY_CELL_LIMIT,
    });
    if (isReadOnly && !decision.allRangesBand) {
      try {
        showToast("Read-only: select an entire row, column, or sheet to change formatting defaults.", "warning");
      } catch {
        // ignore (e.g. toast root missing in tests)
      }
      return;
    }
    if (!decision.allowed) {
      try {
        showToast(
          "Selection is too large to format. Try selecting fewer cells or an entire row/column.",
          "warning",
        );
      } catch {
        // `showToast` requires a #toast-root; ignore failures in test-only contexts.
      }
      return;
    }

    const normalized = rawRanges.map((r) => normalizeSelectionRange(r as Range));
    const expanded = normalized.map((r) => {
      const isFullColBand = r.startRow === 0 && r.endRow === limits.maxRows - 1;
      const isFullRowBand = r.startCol === 0 && r.endCol === limits.maxCols - 1;
      return {
        startRow: r.startRow,
        startCol: r.startCol,
        endRow: isFullColBand ? DEFAULT_GRID_LIMITS.maxRows - 1 : r.endRow,
        endCol: isFullRowBand ? DEFAULT_GRID_LIMITS.maxCols - 1 : r.endCol,
      };
    });

    const useBatch = expanded.length > 1;
    let applied = true;
    if (useBatch) doc.beginBatch({ label: "Format Cells" });
    try {
      for (const r of expanded) {
        const ok = applyFormatCells(
          doc,
          sheetIdNow,
          { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } },
          changes,
        );
        if (ok === false) applied = false;
      }
    } finally {
      if (useBatch) doc.endBatch();
    }
    if (!applied) {
      try {
        showToast("Formatting could not be applied to the full selection. Try selecting fewer cells/rows.", "warning");
      } catch {
        // ignore (e.g. toast root missing in tests)
      }
    }
  }

  applyBtn.addEventListener("click", () => applyFromUi());
  okBtn.addEventListener("click", () => {
    applyFromUi();
    dialog.close("ok");
  });
  cancelBtn.addEventListener("click", () => dialog.close("cancel"));

  dialog.addEventListener("cancel", (e) => {
    e.preventDefault();
    dialog.close("cancel");
  });

  dialog.addEventListener(
    "close",
    () => {
      dialog.remove();
      host.focusGrid();
    },
    { once: true },
  );

  document.body.appendChild(dialog);
  showDialogModal(dialog);
  // If the active style is already a custom number format, focus the code field directly.
  if (numberSelect.value === "__custom__") {
    try {
      numberCustomInput.focus();
    } catch {
      // ignore (e.g. jsdom)
    }
  } else {
    numberSelect.focus();
  }
}
