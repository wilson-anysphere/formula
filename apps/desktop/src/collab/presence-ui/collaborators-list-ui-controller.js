/**
 * @typedef {object} CollaboratorListEntry
 * @property {string} key Stable identifier for diffing (e.g. `${id}:${clientId}`).
 * @property {string} name Display name.
 * @property {string} color CSS color value (e.g. "var(--formula-grid-remote-presence-1)" or "#RRGGBB").
 * @property {string | null | undefined} [sheetName] Optional sheet label (e.g. when viewing a different sheet).
 */

/**
 * Minimal DOM-based "collaborators list" UI.
 *
 * This intentionally avoids React so it can be exercised in unit tests without a bundler,
 * similar to the existing conflicts UI controllers.
 */
export class CollaboratorsListUiController {
  /**
   * @param {object} opts
   * @param {HTMLElement} opts.container Parent element to render into.
   * @param {number | null | undefined} [opts.maxVisible] Optional cap on visible collaborators.
   */
  constructor(opts) {
    if (!opts?.container) throw new Error("CollaboratorsListUiController requires { container }");
    this.container = opts.container;
    this.maxVisible = typeof opts.maxVisible === "number" && Number.isFinite(opts.maxVisible) ? opts.maxVisible : null;

    /** @type {CollaboratorListEntry[]} */
    this._collaborators = [];
    this._signature = "";

    this._rootEl = document.createElement("div");
    this._rootEl.className = "presence-collaborators";
    this._rootEl.dataset.testid = "presence-collaborators";
    this._rootEl.setAttribute("role", "list");
    this._rootEl.setAttribute("aria-label", "Collaborators");
    this.container.appendChild(this._rootEl);
  }

  destroy() {
    this._collaborators = [];
    this._signature = "";
    this._rootEl.remove();
  }

  /**
   * @param {CollaboratorListEntry[]} collaborators
   */
  setCollaborators(collaborators) {
    const next = Array.isArray(collaborators) ? collaborators : [];
    const signature = buildCollaboratorsSignature(next);
    if (signature === this._signature) return;

    this._signature = signature;
    this._collaborators = next;
    this._render();
  }

  _render() {
    const all = this._collaborators;
    if (all.length === 0) {
      this._rootEl.replaceChildren();
      this._rootEl.style.display = "none";
      return;
    }

    this._rootEl.style.display = "flex";

    const maxVisible = this.maxVisible;
    const visible =
      typeof maxVisible === "number" && maxVisible > 0 && all.length > maxVisible ? all.slice(0, maxVisible) : all;
    const overflow =
      typeof maxVisible === "number" && maxVisible > 0 && all.length > maxVisible ? all.length - maxVisible : 0;

    const frag = document.createDocumentFragment();
    for (const entry of visible) {
      const item = document.createElement("div");
      item.className = "presence-collaborators__item";
      item.dataset.testid = "presence-collaborator";
      item.setAttribute("role", "listitem");

      const dot = document.createElement("span");
      dot.className = "presence-collaborators__dot";
      dot.style.backgroundColor = entry.color;
      item.appendChild(dot);

      const name = document.createElement("span");
      name.className = "presence-collaborators__name";
      name.textContent = entry.name;
      item.appendChild(name);

      const sheet = typeof entry.sheetName === "string" ? entry.sheetName.trim() : "";
      if (sheet) {
        const sheetEl = document.createElement("span");
        sheetEl.className = "presence-collaborators__sheet";
        sheetEl.textContent = `• ${sheet}`;
        item.appendChild(sheetEl);
      }

      item.title = sheet ? `${entry.name} • ${sheet}` : entry.name;
      frag.appendChild(item);
    }

    if (overflow > 0) {
      const more = document.createElement("div");
      more.className = "presence-collaborators__overflow";
      more.dataset.testid = "presence-collaborator-overflow";
      more.textContent = `+${overflow}`;
      more.title = `${overflow} more collaborator${overflow === 1 ? "" : "s"}`;
      frag.appendChild(more);
    }

    this._rootEl.replaceChildren(frag);
  }
}

/**
 * @param {CollaboratorListEntry[]} collaborators
 * @returns {string}
 */
function buildCollaboratorsSignature(collaborators) {
  // Keep the signature stable and ignore transient fields so cursor movement does not trigger reflows.
  // Using record-separator boundaries makes it resilient to embedded separators in names/sheet names.
  return collaborators
    .map((c) => {
      const key = typeof c?.key === "string" ? c.key : "";
      const name = typeof c?.name === "string" ? c.name : "";
      const color = typeof c?.color === "string" ? c.color : "";
      const sheet = typeof c?.sheetName === "string" ? c.sheetName : "";
      return `${key}\u001f${name}\u001f${color}\u001f${sheet}`;
    })
    .join("\u001e");
}
