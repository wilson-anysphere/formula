import { cellRefFromKey, numberToCol } from "../../../../../packages/collab/conflicts/src/cell-ref.js";

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
   */
  constructor(opts) {
    this.container = opts.container;
    this.monitor = opts.monitor;

    /** @type {Array<any>} */
    this.conflicts = [];
    /** @type {string | null} */
    this.activeConflictId = null;

    this._toastEl = document.createElement("div");
    this._dialogEl = document.createElement("div");

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
      msg.textContent = `Structural conflict detected (${conflict.type}/${conflict.reason}) in ${formatLocation(conflict)}`;
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
    title.textContent = `Resolve ${conflict.type} conflict (${conflict.reason}) in ${formatLocation(conflict)}`;
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
        label: conflict.remoteUserId ? `Theirs (${conflict.remoteUserId})` : "Theirs",
        conflict,
        side: "remote",
      }),
    );
    dialog.appendChild(body);

    const actions = document.createElement("div");
    actions.className = "conflict-dialog__actions";

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
    pre.textContent = formatJson(summarizeOp(op));
    panel.appendChild(pre);

    if (input.conflict.type === "move") {
      const moveInfo = document.createElement("div");
      moveInfo.dataset.testid = `${input.testid}-move`;
      const toCellKey = op?.toCellKey ?? null;
      if (toCellKey) {
        moveInfo.textContent = `Destination: ${formatCellKey(toCellKey)}`;
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
 */
function formatLocation(conflict) {
  const sheetId = String(conflict?.sheetId ?? "");
  const cell = String(conflict?.cell ?? "");
  if (sheetId && cell) return `${sheetId}!${cell}`;
  return sheetId || cell || "unknown";
}

/**
 * @param {string} cellKey
 */
function formatCellKey(cellKey) {
  try {
    const ref = cellRefFromKey(String(cellKey));
    return `${ref.sheetId}!${numberToCol(ref.col)}${ref.row + 1}`;
  } catch {
    return String(cellKey);
  }
}

/**
 * @param {any} op
 */
function summarizeOp(op) {
  if (!op || typeof op !== "object") return op;
  if (op.kind === "move") {
    return {
      kind: "move",
      from: op.fromCellKey ? formatCellKey(op.fromCellKey) : null,
      to: op.toCellKey ? formatCellKey(op.toCellKey) : null,
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
