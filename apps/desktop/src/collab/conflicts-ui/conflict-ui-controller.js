import { numberToCol } from "../../../../../packages/collab/conflicts/src/cell-ref.js";

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
   * @param {{ resolveConflict: (id: string, formula: string) => boolean }} opts.monitor
   */
  constructor(opts) {
    this.container = opts.container;
    this.monitor = opts.monitor;

    /** @type {Array<any>} */
    this.conflicts = [];
    this.activeConflictId = null;

    this._toastEl = document.createElement("div");
    this._dialogEl = document.createElement("div");

    this._toastEl.dataset.testid = "conflict-toast-root";
    this._dialogEl.dataset.testid = "conflict-dialog-root";

    this.container.appendChild(this._toastEl);
    this.container.appendChild(this._dialogEl);

    this.render();
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
    msg.textContent =
      this.conflicts.length === 1
        ? `Formula conflict detected (${formatCell(this.conflicts[0].cell)})`
        : `${this.conflicts.length} formula conflicts detected`;

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
    title.textContent = `Resolve conflict in ${formatCell(conflict.cell)}`;
    dialog.appendChild(title);

    const body = document.createElement("div");
    body.style.display = "flex";
    body.style.gap = "16px";

    const left = this._renderFormulaPanel({
      testid: "conflict-local",
      label: "Yours",
      formula: conflict.localFormula,
      preview: conflict.localPreview
    });
    const right = this._renderFormulaPanel({
      testid: "conflict-remote",
      label: `Theirs (${conflict.remoteUserId})`,
      formula: conflict.remoteFormula,
      preview: conflict.remotePreview
    });

    body.appendChild(left);
    body.appendChild(right);
    dialog.appendChild(body);

    const actions = document.createElement("div");
    actions.style.display = "flex";
    actions.style.gap = "8px";
    actions.style.marginTop = "12px";

    actions.appendChild(
      this._button("Keep yours", "conflict-choose-local", () => {
        this._resolve(conflict.id, conflict.localFormula);
      })
    );
    actions.appendChild(
      this._button("Use theirs", "conflict-choose-remote", () => {
        this._resolve(conflict.id, conflict.remoteFormula);
      })
    );
    actions.appendChild(
      this._button("Edit…", "conflict-edit", () => {
        this._renderManualEditor(dialog, conflict);
      })
    );
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
    panel.style.flex = "1";

    const label = document.createElement("div");
    label.textContent = input.label;
    label.style.fontWeight = "bold";
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
   * @param {HTMLElement} dialog
   * @param {any} conflict
   */
  _renderManualEditor(dialog, conflict) {
    const existing = dialog.querySelector('[data-testid="conflict-manual-editor"]');
    if (existing) existing.remove();

    const editor = document.createElement("div");
    editor.dataset.testid = "conflict-manual-editor";
    editor.style.marginTop = "12px";

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
   * @param {string} chosenFormula
   */
  _resolve(conflictId, chosenFormula) {
    const ok = this.monitor.resolveConflict(conflictId, chosenFormula);
    if (!ok) return;

    this.conflicts = this.conflicts.filter((c) => c.id !== conflictId);
    this.activeConflictId = this.conflicts[0]?.id ?? null;
    this.render();
  }
}

/**
 * @param {{ sheetId: string, row: number, col: number }} cell
 */
function formatCell(cell) {
  return `${cell.sheetId}!${numberToCol(cell.col)}${cell.row + 1}`;
}
