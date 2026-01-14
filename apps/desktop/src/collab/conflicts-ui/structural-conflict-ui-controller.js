import { cellRefFromKey, numberToCol } from "../../../../../packages/collab/conflicts/src/cell-ref.js";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.ts";
import { markKeybindingBarrier } from "../../keybindingBarrier.js";
import { renderFormulaDiffDom } from "../../versioning/ui/formulaDiffDom.ts";

/**
 * A minimal DOM-based UI for resolving *structural* (move/delete-vs-edit) cell
 * conflicts.
 *
 * This intentionally avoids React so it can be exercised in tests without a
 * bundler or app shell wiring.
 */
export class StructuralConflictUiController {
  /**
   * @param {object} opts
   * @param {HTMLElement} opts.container
   * @param {{ resolveConflict: (id: string, resolution: any) => boolean }} opts.monitor
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
    /** @type {string | null} */
    this.activeConflictId = null;

    this._toastEl = document.createElement("div");
    this._dialogEl = document.createElement("div");
    markKeybindingBarrier(this._dialogEl);

    this._toastEl.dataset.testid = "structural-conflict-toast-root";
    this._dialogEl.dataset.testid = "structural-conflict-dialog-root";

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
    if (!conflict) return;
    this.conflicts.push(conflict);
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
    toast.dataset.testid = "structural-conflict-toast";

    const msg = document.createElement("div");
    if (this.conflicts.length === 1) {
      const conflict = this.conflicts[0];
      msg.textContent = `Structural conflict detected (${conflict.type}/${conflict.reason}) in ${formatLocation(
        conflict,
        this.sheetNameResolver,
      )}`;
    } else {
      msg.textContent = `${this.conflicts.length} structural conflicts detected`;
    }

    const btn = document.createElement("button");
    btn.textContent = "Resolve…";
    btn.dataset.testid = "structural-conflict-toast-open";
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
    dialog.dataset.testid = "structural-conflict-dialog";

    const title = document.createElement("h2");
    title.textContent = `Resolve ${conflict.type} conflict (${conflict.reason}) in ${formatLocation(
      conflict,
      this.sheetNameResolver,
    )}`;
    dialog.appendChild(title);

    const meta = document.createElement("div");
    meta.dataset.testid = "structural-conflict-meta";
    meta.textContent = `type: ${String(conflict.type)} | reason: ${String(conflict.reason)}`;
    dialog.appendChild(meta);

    const body = document.createElement("div");
    body.className = "conflict-dialog__body";

    body.appendChild(
      this._renderSide({
        testid: "structural-conflict-local",
        label: "Yours",
        conflict,
        side: "local",
      }),
    );
    body.appendChild(
      this._renderSide({
        testid: "structural-conflict-remote",
        label: formatRemoteLabel(conflict.remoteUserId, this.resolveUserLabel),
        conflict,
        side: "remote",
      }),
    );
    dialog.appendChild(body);

    const localFormula = extractFormulaFromOp(conflict.local);
    const remoteFormula = extractFormulaFromOp(conflict.remote);
    if (localFormula !== null || remoteFormula !== null) {
      dialog.appendChild(
        renderFormulaDiffDom(localFormula, remoteFormula, {
          testid: "structural-conflict-formula-diff",
          label: "Formula diff",
        }),
      );
    }

    const actions = document.createElement("div");
    actions.className = "conflict-dialog__actions";

    actions.appendChild(
      this._button("Jump to cell", "structural-conflict-jump-to-cell", () => {
        if (!this.onNavigateToCell) return;
        const cellKey = String(conflict?.cellKey ?? "");
        if (!cellKey) return;
        try {
          const ref = cellRefFromKey(cellKey);
          const sheetId = String(ref?.sheetId ?? "");
          const row = Number(ref?.row);
          const col = Number(ref?.col);
          if (!sheetId) return;
          if (!Number.isInteger(row) || row < 0) return;
          if (!Number.isInteger(col) || col < 0) return;
          this.onNavigateToCell({ sheetId, row, col });
        } catch {
          // Best-effort; ignore invalid keys.
        }
      }),
    );

    if (conflict.type === "move") {
      actions.appendChild(
        this._button("Keep ours destination", "structural-conflict-choose-ours", () => {
          this._resolve(conflict.id, { choice: "ours" });
        }),
      );
      actions.appendChild(
        this._button("Use theirs destination", "structural-conflict-choose-theirs", () => {
          this._resolve(conflict.id, { choice: "theirs" });
        }),
      );
    } else {
      actions.appendChild(
        this._button("Keep ours", "structural-conflict-choose-ours", () => {
          this._resolve(conflict.id, { choice: "ours" });
        }),
      );
      actions.appendChild(
        this._button("Use theirs", "structural-conflict-choose-theirs", () => {
          this._resolve(conflict.id, { choice: "theirs" });
        }),
      );
    }

    actions.appendChild(
      this._button("Manual…", "structural-conflict-manual", () => {
        this._renderManualEditor(dialog, conflict);
      }),
    );

    actions.appendChild(
      this._button("Close", "structural-conflict-close", () => {
        this.activeConflictId = null;
        this.render();
      }),
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
  _renderSide(input) {
    const panel = document.createElement("div");
    panel.dataset.testid = input.testid;
    panel.className = "conflict-dialog__panel";

    const label = document.createElement("div");
    label.textContent = input.label;
    label.className = "conflict-dialog__panel-label";
    panel.appendChild(label);

    const pre = document.createElement("pre");
    const op = input.side === "local" ? input.conflict.local : input.conflict.remote;
    pre.textContent = formatJson(summarizeOp(op, this.sheetNameResolver));
    panel.appendChild(pre);

    if (input.conflict.type === "move") {
      const moveInfo = document.createElement("div");
      moveInfo.dataset.testid = `${input.testid}-move`;
      const toCellKey = op?.toCellKey ?? null;
      if (toCellKey) {
        moveInfo.textContent = `Destination: ${formatCellKey(toCellKey, this.sheetNameResolver)}`;
      } else {
        moveInfo.textContent = "";
      }
      panel.appendChild(moveInfo);
    }

    return panel;
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
   * @param {HTMLElement} dialog
   * @param {any} conflict
   */
  _renderManualEditor(dialog, conflict) {
    const existing = dialog.querySelector('[data-testid="structural-conflict-manual-editor"]');
    if (existing) existing.remove();

    const editor = document.createElement("div");
    editor.dataset.testid = "structural-conflict-manual-editor";
    editor.className = "conflict-dialog__manual-editor";

    if (conflict.type === "move") {
      const destLabel = document.createElement("div");
      destLabel.textContent = "Destination cellKey (e.g. Sheet1:0:1)";
      editor.appendChild(destLabel);

      const destInput = document.createElement("input");
      destInput.type = "text";
      destInput.size = 40;
      destInput.dataset.testid = "structural-conflict-manual-destination";
      destInput.value = conflict.local?.toCellKey ?? conflict.remote?.toCellKey ?? "";
      editor.appendChild(destInput);

       const cellLabel = document.createElement("div");
       cellLabel.className = "conflict-dialog__manual-label";
       cellLabel.textContent = "Optional moved cell JSON (leave blank to use ours/theirs content)";
       editor.appendChild(cellLabel);

      const textarea = document.createElement("textarea");
      textarea.rows = 5;
      textarea.cols = 60;
      textarea.dataset.testid = "structural-conflict-manual-cell";
      textarea.value = formatJson(conflict.local?.cell ?? null);
      editor.appendChild(textarea);

      const apply = this._button("Apply", "structural-conflict-manual-apply", () => {
        const to = destInput.value.trim();
        if (!to) return;
        const raw = textarea.value.trim();
        let cell = undefined;
        if (raw) {
          try {
            cell = JSON.parse(raw);
          } catch {
            return;
          }
        }
        const resolution = { choice: "manual", to, ...(cell !== undefined ? { cell } : {}) };
        this._resolve(conflict.id, resolution);
      });
      editor.appendChild(apply);

      dialog.appendChild(editor);
      return;
    }

    const label = document.createElement("div");
    label.textContent = "Manual cell JSON (leave blank for delete)";
    editor.appendChild(label);

    const textarea = document.createElement("textarea");
    textarea.rows = 6;
    textarea.cols = 60;
    textarea.dataset.testid = "structural-conflict-manual-cell";
    textarea.value = formatJson(conflict.local?.after ?? null);
    editor.appendChild(textarea);

    const apply = this._button("Apply", "structural-conflict-manual-apply", () => {
      const raw = textarea.value.trim();
      if (!raw) {
        this._resolve(conflict.id, { choice: "manual", cell: null });
        return;
      }
      let cell;
      try {
        cell = JSON.parse(raw);
      } catch {
        return;
      }
      this._resolve(conflict.id, { choice: "manual", cell });
    });

    editor.appendChild(apply);
    dialog.appendChild(editor);
  }

  /**
   * @param {string} conflictId
   * @param {any} resolution
   */
  _resolve(conflictId, resolution) {
    const ok = this.monitor.resolveConflict(conflictId, resolution);
    if (!ok) return;

    this.conflicts = this.conflicts.filter((c) => c.id !== conflictId);
    if (this.conflicts.length === 0) {
      this.activeConflictId = null;
    } else if (this.activeConflictId === conflictId) {
      this.activeConflictId = this.conflicts[0]?.id ?? null;
    }
    this.render();
  }
}

/**
 * @param {any} conflict
 * @param {import("../../sheet/sheetNameResolver.ts").SheetNameResolver | null} sheetNameResolver
 */
function formatLocation(conflict, sheetNameResolver) {
  const sheetId = String(conflict?.sheetId ?? "");
  const cell = String(conflict?.cell ?? "");
  const sheetName = sheetNameResolver?.getSheetNameById?.(sheetId) ?? sheetId;
  const sheetPrefix = sheetName ? `${formatSheetNameForA1(sheetName)}!` : "";
  if (sheetPrefix && cell) return `${sheetPrefix}${cell}`;
  return sheetName || cell || "unknown";
}

/**
 * @param {string} cellKey
 * @param {import("../../sheet/sheetNameResolver.ts").SheetNameResolver | null} sheetNameResolver
 */
function formatCellKey(cellKey, sheetNameResolver) {
  try {
    const ref = cellRefFromKey(String(cellKey));
    const sheetName = sheetNameResolver?.getSheetNameById?.(ref.sheetId) ?? ref.sheetId;
    return `${formatSheetNameForA1(sheetName)}!${numberToCol(ref.col)}${ref.row + 1}`;
  } catch {
    return String(cellKey);
  }
}

/**
 * @param {any} op
 * @param {import("../../sheet/sheetNameResolver.ts").SheetNameResolver | null} sheetNameResolver
 */
function summarizeOp(op, sheetNameResolver) {
  if (!op || typeof op !== "object") return op;
  if (op.kind === "move") {
    return {
      kind: "move",
      from: op.fromCellKey ? formatCellKey(op.fromCellKey, sheetNameResolver) : null,
      to: op.toCellKey ? formatCellKey(op.toCellKey, sheetNameResolver) : null,
      cell: op.cell ?? null,
    };
  }
  if (op.kind === "edit" || op.kind === "delete") {
    return {
      kind: op.kind,
      before: "before" in op ? op.before ?? null : null,
      after: "after" in op ? op.after ?? null : null,
    };
  }
  return op;
}

/**
 * Format a sheet name token for A1 references using Excel quoting conventions.
 * (Only used for display in the conflict UI; underlying sheet ids remain stable.)
 *
 * @param {string} sheetName
 */
/**
 * @param {any} value
 */
function formatJson(value) {
  if (value === null) return "null";
  if (value === undefined) return "undefined";
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

/**
 * @param {string} remoteUserId
 * @param {((userId: string) => string) | null} resolver
 */
function formatRemoteLabel(remoteUserId, resolver) {
  const id = String(remoteUserId ?? "");
  if (!id) return "Theirs";

  let resolved = "";
  try {
    resolved = resolver ? resolver(id) : "";
  } catch {
    resolved = "";
  }

  const trimmed = typeof resolved === "string" ? resolved.trim() : "";
  const label = trimmed ? trimmed : id;
  return `Theirs (${label})`;
}

/**
 * @param {any} op
 * @returns {string | null}
 */
function extractFormulaFromOp(op) {
  if (!op || typeof op !== "object") return null;
  let cell = null;
  if (op.kind === "move") cell = op.cell ?? null;
  else if (op.kind === "edit" || op.kind === "delete") cell = op.after ?? null;
  if (!cell || typeof cell !== "object") return null;
  const formula = cell.formula;
  return typeof formula === "string" && formula.trim() ? formula : null;
}
