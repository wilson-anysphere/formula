import { numberToCol } from "../../../../../packages/collab/conflicts/src/cell-ref.js";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.ts";
import { markKeybindingBarrier } from "../../keybindingBarrier.js";
import { renderFormulaDiffDom } from "../../versioning/ui/formulaDiffDom.ts";

/**
 * A minimal DOM-based conflict UX for the desktop app.
 *
 * The real application will likely use React, but this implementation is fully
 * functional and easy to exercise in tests without a bundler.
 */
export class ConflictUiController {
  /**
   * @param {object} opts
   * @param {HTMLElement} opts.container
   * @param {{ resolveConflict: (id: string, chosen: any) => boolean }} opts.monitor
   * @param {import("../../sheet/sheetNameResolver.ts").SheetNameResolver | null | undefined} [opts.sheetNameResolver]
   * @param {(cellRef: { sheetId: string, row: number, col: number }) => void} [opts.onNavigateToCell]
   * @param {(userId: string) => string} [opts.resolveUserLabel]
   */
  constructor(opts) {
    this.container = opts.container;
    this.monitor = opts.monitor;
    this.sheetNameResolver = opts.sheetNameResolver ?? null;
    this.onNavigateToCell = typeof opts.onNavigateToCell === "function" ? opts.onNavigateToCell : null;
    this.resolveUserLabel = typeof opts.resolveUserLabel === "function" ? opts.resolveUserLabel : null;

    /** @type {Array<any>} */
    this.conflicts = [];
    this.activeConflictId = null;

    this._toastEl = document.createElement("div");
    this._dialogEl = document.createElement("div");
    markKeybindingBarrier(this._dialogEl);

    this._toastEl.dataset.testid = "conflict-toast-root";
    this._dialogEl.dataset.testid = "conflict-dialog-root";

    this.container.appendChild(this._toastEl);
    this.container.appendChild(this._dialogEl);

    this.render();
  }

  destroy() {
    this.conflicts = [];
    this.activeConflictId = null;
    this._toastEl.remove();
    this._dialogEl.remove();
  }

  /**
   * @param {any} conflict
   */
  addConflict(conflict) {
    this.conflicts.push(conflict);
    if (!this.activeConflictId) this.activeConflictId = conflict.id;
    this.render();
  }

  render() {
    this._renderToast();
    this._renderDialog();
  }

  _renderToast() {
    this._toastEl.innerHTML = "";
    if (this.conflicts.length === 0) return;

    const toast = document.createElement("div");
    toast.dataset.testid = "conflict-toast";

    const msg = document.createElement("div");
    if (this.conflicts.length === 1) {
      const conflict = this.conflicts[0];
      const label = conflict.kind === "value" ? "Value" : conflict.kind === "content" ? "Content" : "Formula";
      msg.textContent = `${label} conflict detected (${formatCell(conflict.cell, this.sheetNameResolver)})`;
    } else {
      msg.textContent = `${this.conflicts.length} conflicts detected`;
    }

    const btn = document.createElement("button");
    btn.textContent = "Resolve…";
    btn.dataset.testid = "conflict-toast-open";
    btn.addEventListener("click", () => {
      this.activeConflictId = this.conflicts[0]?.id ?? null;
      this.render();
    });

    toast.appendChild(msg);
    toast.appendChild(btn);
    this._toastEl.appendChild(toast);
  }

  _renderDialog() {
    this._dialogEl.innerHTML = "";
    const conflict = this.conflicts.find((c) => c.id === this.activeConflictId) ?? null;
    if (!conflict) return;

    const dialog = document.createElement("div");
    dialog.dataset.testid = "conflict-dialog";

    const title = document.createElement("h2");
    const label = conflict.kind === "value" ? "value" : conflict.kind === "content" ? "content" : "formula";
    title.textContent = `Resolve ${label} conflict in ${formatCell(conflict.cell, this.sheetNameResolver)}`;
    dialog.appendChild(title);

    const body = document.createElement("div");
    body.className = "conflict-dialog__body";

    const left = this._renderConflictSide({
      testid: "conflict-local",
      label: "Yours",
      conflict,
      side: "local"
    });

    let resolvedRemote = "";
    if (conflict.remoteUserId && this.resolveUserLabel) {
      try {
        resolvedRemote = this.resolveUserLabel(conflict.remoteUserId);
      } catch {
        resolvedRemote = "";
      }
    }
    const resolvedRemoteTrimmed = typeof resolvedRemote === "string" ? resolvedRemote.trim() : "";
    const remoteLabel = conflict.remoteUserId && resolvedRemoteTrimmed ? resolvedRemoteTrimmed : conflict.remoteUserId;
    const right = this._renderConflictSide({
      testid: "conflict-remote",
      label: remoteLabel ? `Theirs (${remoteLabel})` : "Theirs",
      conflict,
      side: "remote"
    });

    body.appendChild(left);
    body.appendChild(right);
    dialog.appendChild(body);

    const maybeFormulaDiff =
      conflict.kind === "formula"
        ? { before: conflict.localFormula ?? null, after: conflict.remoteFormula ?? null }
        : conflict.kind === "content" && (conflict.local?.type === "formula" || conflict.remote?.type === "formula")
          ? {
              before: conflict.local?.type === "formula" ? conflict.local.formula ?? null : null,
              after: conflict.remote?.type === "formula" ? conflict.remote.formula ?? null : null,
            }
          : null;
    if (maybeFormulaDiff) {
      dialog.appendChild(renderFormulaDiffDom(maybeFormulaDiff.before, maybeFormulaDiff.after, { testid: "conflict-formula-diff" }));
    }

    const actions = document.createElement("div");
    actions.className = "conflict-dialog__actions";

    actions.appendChild(
      this._button("Jump to cell", "conflict-jump-to-cell", () => {
        if (!this.onNavigateToCell) return;
        const cell = conflict?.cell;
        if (!cell || typeof cell !== "object") return;
        const sheetId = String(cell.sheetId ?? "").trim();
        const row = Number(cell.row);
        const col = Number(cell.col);
        if (!sheetId) return;
        if (!Number.isInteger(row) || row < 0) return;
        if (!Number.isInteger(col) || col < 0) return;
        try {
          this.onNavigateToCell({ sheetId, row, col });
        } catch {
          // Best-effort: ignore navigation failures so conflict UI remains usable.
        }
      })
    );

    actions.appendChild(
      this._button("Keep yours", "conflict-choose-local", () => {
        const chosen =
          conflict.kind === "content" ? conflict.local : conflict.kind === "value" ? conflict.localValue : conflict.localFormula;
        this._resolve(conflict.id, chosen);
      })
    );
    actions.appendChild(
      this._button("Use theirs", "conflict-choose-remote", () => {
        const chosen =
          conflict.kind === "content" ? conflict.remote : conflict.kind === "value" ? conflict.remoteValue : conflict.remoteFormula;
        this._resolve(conflict.id, chosen);
      })
    );
    if (conflict.kind === "formula") {
      actions.appendChild(
        this._button("Edit…", "conflict-edit", () => {
          this._renderManualEditor(dialog, conflict);
        })
      );
    }
    actions.appendChild(
      this._button("Close", "conflict-close", () => {
        this.activeConflictId = null;
        this.render();
      })
    );

    dialog.appendChild(actions);
    this._dialogEl.appendChild(dialog);
  }

  /**
   * @param {object} input
   * @param {string} input.testid
   * @param {string} input.label
   * @param {any} input.conflict
   * @param {"local" | "remote"} input.side
   */
  _renderConflictSide(input) {
    const { conflict } = input;

    if (conflict.kind === "content") {
      const choice = input.side === "local" ? conflict.local : conflict.remote;
      if (choice?.type === "formula") {
        return this._renderFormulaPanel({
          testid: input.testid,
          label: input.label,
          formula: choice.formula,
          preview: choice.preview
        });
      }
      return this._renderValuePanel({ testid: input.testid, label: input.label, value: choice?.value ?? null });
    }

    if (conflict.kind === "value") {
      return this._renderValuePanel({
        testid: input.testid,
        label: input.label,
        value: input.side === "local" ? conflict.localValue : conflict.remoteValue
      });
    }

    return this._renderFormulaPanel({
      testid: input.testid,
      label: input.label,
      formula: input.side === "local" ? conflict.localFormula : conflict.remoteFormula,
      preview: input.side === "local" ? conflict.localPreview : conflict.remotePreview
    });
  }

  /**
   * @param {string} label
   * @param {string} testid
   * @param {() => void} onClick
   */
  _button(label, testid, onClick) {
    const btn = document.createElement("button");
    btn.textContent = label;
    btn.dataset.testid = testid;
    btn.addEventListener("click", onClick);
    return btn;
  }

  /**
   * @param {object} input
   * @param {string} input.testid
   * @param {string} input.label
   * @param {string} input.formula
   * @param {any} input.preview
   */
  _renderFormulaPanel(input) {
    const panel = document.createElement("div");
    panel.dataset.testid = input.testid;
    panel.className = "conflict-dialog__panel";

    const label = document.createElement("div");
    label.textContent = input.label;
    label.className = "conflict-dialog__panel-label";
    panel.appendChild(label);

    const pre = document.createElement("pre");
    pre.textContent = input.formula;
    panel.appendChild(pre);

    const preview = document.createElement("div");
    preview.dataset.testid = `${input.testid}-preview`;
    preview.textContent =
      input.preview === undefined ? "" : `Preview: ${input.preview === null ? "—" : String(input.preview)}`;
    panel.appendChild(preview);

    return panel;
  }

  /**
   * @param {object} input
   * @param {string} input.testid
   * @param {string} input.label
   * @param {any} input.value
   */
  _renderValuePanel(input) {
    const panel = document.createElement("div");
    panel.dataset.testid = input.testid;
    panel.className = "conflict-dialog__panel";

    const label = document.createElement("div");
    label.textContent = input.label;
    label.className = "conflict-dialog__panel-label";
    panel.appendChild(label);

    const pre = document.createElement("pre");
    pre.textContent = formatValue(input.value);
    panel.appendChild(pre);

    return panel;
  }

  /**
   * @param {HTMLElement} dialog
   * @param {any} conflict
   */
  _renderManualEditor(dialog, conflict) {
    const existing = dialog.querySelector('[data-testid="conflict-manual-editor"]');
    if (existing) existing.remove();

    const editor = document.createElement("div");
    editor.dataset.testid = "conflict-manual-editor";
    editor.className = "conflict-dialog__manual-editor";

    const textarea = document.createElement("textarea");
    textarea.rows = 4;
    textarea.cols = 60;
    textarea.value = conflict.localFormula;
    textarea.dataset.testid = "conflict-manual-textarea";

    const apply = this._button("Apply", "conflict-manual-apply", () => {
      this._resolve(conflict.id, textarea.value);
    });

    editor.appendChild(textarea);
    editor.appendChild(apply);
    dialog.appendChild(editor);
  }

  /**
   * @param {string} conflictId
   * @param {any} chosen
   */
  _resolve(conflictId, chosen) {
    const ok = this.monitor.resolveConflict(conflictId, chosen);
    if (!ok) return;

    this.conflicts = this.conflicts.filter((c) => c.id !== conflictId);
    this.activeConflictId = this.conflicts[0]?.id ?? null;
    this.render();
  }
}

/**
 * @param {{ sheetId: string, row: number, col: number }} cell
 * @param {import("../../sheet/sheetNameResolver.ts").SheetNameResolver | null} sheetNameResolver
 */
function formatCell(cell, sheetNameResolver) {
  const sheetId = String(cell?.sheetId ?? "");
  const sheetName = sheetNameResolver?.getSheetNameById?.(sheetId) ?? sheetId;
  return `${formatSheetNameForA1(sheetName)}!${numberToCol(cell.col)}${cell.row + 1}`;
}

/**
 * @param {any} value
 */
function formatValue(value) {
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  if (typeof value === "string") return JSON.stringify(value);
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}
