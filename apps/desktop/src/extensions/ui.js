/**
 * Extension UI helpers (toasts + simple modal prompts).
 *
 * This file is intentionally plain `.js` so it can be imported by node:test suites
 * without requiring a TypeScript runtime.
 *
 * @typedef {"info" | "warning" | "error"} MessageType
 */

let extensionsUiDialogTitleId = 0;
/**
 * @param {"input-box" | "quick-pick"} kind
 * @returns {string}
 */
function nextExtensionsUiDialogTitleId(kind) {
  extensionsUiDialogTitleId += 1;
  return `extensions-ui-${kind}-title-${extensionsUiDialogTitleId}`;
}

/**
 * @param {HTMLDialogElement} dialog
 */
function showModal(dialog) {
  // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
  if (typeof dialog.showModal === "function") {
    // @ts-expect-error - HTMLDialogElement.showModal() not implemented in jsdom.
    dialog.showModal();
    return;
  }
  // jsdom doesn't implement showModal(). Best-effort fallback so unit tests can
  // drive the DOM without a full dialog polyfill.
  dialog.setAttribute("open", "true");
}

/**
 * @param {HTMLDialogElement} dialog
 * @returns {boolean}
 */
function isDialogOpen(dialog) {
  // @ts-expect-error - jsdom typing mismatch.
  return dialog.open === true || dialog.hasAttribute("open");
}

/**
 * @param {HTMLDialogElement} dialog
 * @param {string} returnValue
 */
function closeDialog(dialog, returnValue) {
  if (!isDialogOpen(dialog)) return;
  // @ts-expect-error - HTMLDialogElement.close() not implemented in jsdom.
  if (typeof dialog.close === "function") {
    // @ts-expect-error - HTMLDialogElement.close() not implemented in jsdom.
    dialog.close(returnValue);
    return;
  }
  // jsdom doesn't implement close()/returnValue. Emulate the close contract so
  // our prompt logic and tests can rely on the close event.
  // @ts-expect-error - returnValue not modeled on jsdom's dialog typings.
  dialog.returnValue = returnValue;
  dialog.removeAttribute("open");
  dialog.dispatchEvent(new Event("close"));
}

/**
 * @returns {HTMLElement}
 */
function ensureToastRoot() {
  const root = document.getElementById("toast-root");
  if (!root) {
    throw new Error("Missing #toast-root (required for extension UI messages)");
  }
  return root;
}

/**
 * @param {string} message
 * @param {MessageType} [type]
 * @param {{ timeoutMs?: number }} [options]
 */
export function showToast(message, type = "info", options = {}) {
  const root = ensureToastRoot();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.dataset.type = type;
  toast.dataset.testid = "toast";
  toast.textContent = message;

  root.appendChild(toast);

  const timeoutMs = Number.isFinite(options.timeoutMs) ? Math.max(0, options.timeoutMs ?? 0) : 4_000;
  const timeout = window.setTimeout(() => {
    toast.remove();
  }, timeoutMs);

  toast.addEventListener("click", () => {
    window.clearTimeout(timeout);
    toast.remove();
  });
}

/**
 * @typedef {{
 *   prompt?: string;
 *   value?: string;
 *   placeHolder?: string;
 *   type?: "text" | "password" | "textarea";
 *   rows?: number;
 *   okLabel?: string;
 *   cancelLabel?: string;
 * }} InputBoxOptions
 */

/**
 * @param {InputBoxOptions} [options]
 * @returns {Promise<string | null>}
 */
export async function showInputBox(options = {}) {
  if (typeof document === "undefined" || !document.body) return null;

  const dialog = document.createElement("dialog");
  dialog.className = "dialog extensions-ui";
  dialog.dataset.testid = "input-box";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = options.prompt ?? "Input";
  title.id = nextExtensionsUiDialogTitleId("input-box");
  dialog.setAttribute("aria-labelledby", title.id);

  const mode = options.type ?? "text";
  /** @type {HTMLInputElement | HTMLTextAreaElement} */
  const field = (() => {
    if (mode === "textarea") {
      const textarea = document.createElement("textarea");
      textarea.className = "dialog__field";
      textarea.value = options.value ?? "";
      if (options.placeHolder) textarea.placeholder = options.placeHolder;
      textarea.rows = Math.max(3, options.rows ?? 10);
      // Style textarea via CSS (see styles/extensions-ui.css); avoid inline styles so the
      // dialog stays theme/token driven and testable via node/jsdom.
      // The dialog title doubles as the textarea label.
      textarea.setAttribute("aria-labelledby", title.id);
      textarea.dataset.testid = "input-box-field";
      return textarea;
    }

    const input = document.createElement("input");
    input.type = mode === "password" ? "password" : "text";
    input.className = "dialog__field";
    input.value = options.value ?? "";
    if (options.placeHolder) input.placeholder = options.placeHolder;
    // The dialog title doubles as the input label.
    input.setAttribute("aria-labelledby", title.id);
    input.dataset.testid = "input-box-field";
    return input;
  })();

  const controls = document.createElement("div");
  controls.className = "dialog__controls";

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = options.cancelLabel ?? "Cancel";
  cancelBtn.dataset.testid = "input-box-cancel";

  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.textContent = options.okLabel ?? "OK";
  okBtn.dataset.testid = "input-box-ok";

  controls.appendChild(cancelBtn);
  controls.appendChild(okBtn);

  dialog.appendChild(title);
  dialog.appendChild(field);
  dialog.appendChild(controls);

  document.body.appendChild(dialog);

  const cleanup = () => {
    dialog.remove();
  };

  return await new Promise((resolve) => {
    dialog.addEventListener(
      "close",
      () => {
        // @ts-expect-error - returnValue not modeled on jsdom's dialog typings.
        const returnValue = String(dialog.returnValue ?? "");
        cleanup();
        resolve(returnValue === "ok" ? field.value : null);
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      closeDialog(dialog, "cancel");
    });

    cancelBtn.addEventListener("click", () => closeDialog(dialog, "cancel"));
    okBtn.addEventListener("click", () => closeDialog(dialog, "ok"));

    showModal(dialog);
    field.focus();

    field.addEventListener("keydown", (e) => {
      if (mode === "textarea") {
        // Let enter insert a newline; allow Ctrl/Cmd+Enter to submit.
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
          e.preventDefault();
          closeDialog(dialog, "ok");
        }
        return;
      }

      if (e.key === "Enter") {
        e.preventDefault();
        closeDialog(dialog, "ok");
      } else if (e.key === "Escape") {
        e.preventDefault();
        closeDialog(dialog, "cancel");
      }
    });
  });
}

/**
 * @template T
 * @typedef {{ label: string; value: T; description?: string; detail?: string }} QuickPickItem
 */

/**
 * @template T
 * @param {QuickPickItem<T>[]} items
 * @param {{ placeHolder?: string }} [options]
 * @returns {Promise<T | null>}
 */
export async function showQuickPick(items, options = {}) {
  if (typeof document === "undefined" || !document.body) return null;

  const dialog = document.createElement("dialog");
  dialog.className = "dialog extensions-ui";
  dialog.dataset.testid = "quick-pick";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = options.placeHolder ?? "Select an item";
  title.id = nextExtensionsUiDialogTitleId("quick-pick");
  dialog.setAttribute("aria-labelledby", title.id);

  const list = document.createElement("div");
  list.className = "quick-pick__list";

  dialog.appendChild(title);
  dialog.appendChild(list);

  document.body.appendChild(dialog);

  const cleanup = () => {
    dialog.remove();
  };

  return await new Promise((resolve) => {
    /** @type {HTMLButtonElement[]} */
    const buttons = [];

    dialog.addEventListener(
      "close",
      () => {
        // @ts-expect-error - returnValue not modeled on jsdom's dialog typings.
        const returnValue = String(dialog.returnValue ?? "");
        cleanup();
        if (!returnValue) resolve(null);
        else resolve(items[Number(returnValue)]?.value ?? null);
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      closeDialog(dialog, "");
    });

    /** @type {HTMLButtonElement | null} */
    let firstBtn = null;
    for (const [idx, item] of items.entries()) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "quick-pick__item";
      btn.dataset.testid = `quick-pick-item-${idx}`;
      if (!firstBtn) firstBtn = btn;
      buttons.push(btn);

      const label = document.createElement("div");
      label.textContent = item.label;
      label.className = "quick-pick__label";
      btn.appendChild(label);

      const secondary = item.description ?? item.detail ?? null;
      if (secondary) {
        const sub = document.createElement("div");
        sub.textContent = secondary;
        sub.className = "quick-pick__subtext";
        btn.appendChild(sub);
      }

      btn.addEventListener("click", () => {
        closeDialog(dialog, String(idx));
      });

      list.appendChild(btn);
    }

    dialog.addEventListener("keydown", (e) => {
      if (e.defaultPrevented) return;

      // Arrow-key navigation between list items (VS Code-like quick pick behavior).
      if (buttons.length === 0) return;

      if (e.key === "Escape") {
        e.preventDefault();
        closeDialog(dialog, "");
        return;
      }
      const active = document.activeElement;
      const currentIndex = active ? buttons.indexOf(/** @type {HTMLButtonElement} */ (active)) : -1;

      /**
       * @param {number} idx
       */
      const focusIndex = (idx) => {
        const clamped = Math.max(0, Math.min(idx, buttons.length - 1));
        const btn = buttons[clamped];
        if (!btn) return;
        try {
          btn.focus();
          btn.scrollIntoView({ block: "nearest" });
        } catch {
          // Best-effort focus/scroll.
        }
      };

      if (e.key === "ArrowDown") {
        e.preventDefault();
        focusIndex(currentIndex === -1 ? 0 : currentIndex + 1);
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        focusIndex(currentIndex <= 0 ? 0 : currentIndex - 1);
      } else if (e.key === "Home") {
        e.preventDefault();
        focusIndex(0);
      } else if (e.key === "End") {
        e.preventDefault();
        focusIndex(buttons.length - 1);
      }
    });

    showModal(dialog);
    try {
      firstBtn?.focus();
    } catch {
      // Best-effort focus.
    }
  });
}
