export type MessageType = "info" | "warning" | "error";

let extensionsUiDialogTitleId = 0;
function nextExtensionsUiDialogTitleId(kind: "input-box" | "quick-pick"): string {
  extensionsUiDialogTitleId += 1;
  return `extensions-ui-${kind}-title-${extensionsUiDialogTitleId}`;
}

function showModal(dialog: HTMLDialogElement): void {
  if (typeof (dialog as any).showModal === "function") {
    (dialog as any).showModal();
    return;
  }
  // jsdom doesn't implement showModal(). Best-effort fallback so unit tests can
  // drive the DOM without a full dialog polyfill.
  dialog.setAttribute("open", "true");
}

function closeDialog(dialog: HTMLDialogElement, returnValue: string): void {
  if (typeof (dialog as any).close === "function") {
    (dialog as any).close(returnValue);
    return;
  }
  // jsdom doesn't implement close()/returnValue. Emulate the close contract so
  // our prompt logic and tests can rely on the close event.
  (dialog as any).returnValue = returnValue;
  dialog.removeAttribute("open");
  dialog.dispatchEvent(new Event("close"));
}

function ensureToastRoot(): HTMLElement {
  const root = document.getElementById("toast-root");
  if (!root) {
    throw new Error("Missing #toast-root (required for extension UI messages)");
  }
  return root;
}

export function showToast(message: string, type: MessageType = "info", options: { timeoutMs?: number } = {}): void {
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

type InputBoxOptions = { prompt?: string; value?: string; placeHolder?: string };

export async function showInputBox(options: InputBoxOptions = {}): Promise<string | null> {
  const dialog = document.createElement("dialog");
  dialog.className = "dialog extensions-ui";
  dialog.dataset.testid = "input-box";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = options.prompt ?? "Input";
  title.id = nextExtensionsUiDialogTitleId("input-box");
  dialog.setAttribute("aria-labelledby", title.id);

  const input = document.createElement("input");
  input.type = "text";
  input.className = "dialog__field";
  input.value = options.value ?? "";
  if (options.placeHolder) input.placeholder = options.placeHolder;
  // The dialog title doubles as the input label.
  input.setAttribute("aria-labelledby", title.id);
  input.dataset.testid = "input-box-field";

  const controls = document.createElement("div");
  controls.className = "dialog__controls";

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = "Cancel";
  cancelBtn.dataset.testid = "input-box-cancel";

  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.textContent = "OK";
  okBtn.dataset.testid = "input-box-ok";

  controls.appendChild(cancelBtn);
  controls.appendChild(okBtn);

  dialog.appendChild(title);
  dialog.appendChild(input);
  dialog.appendChild(controls);

  document.body.appendChild(dialog);

  const cleanup = () => {
    dialog.remove();
  };

  return await new Promise<string | null>((resolve) => {
    dialog.addEventListener(
      "close",
      () => {
        const returnValue = String((dialog as any).returnValue ?? dialog.returnValue ?? "");
        cleanup();
        resolve(returnValue === "ok" ? input.value : null);
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
    input.focus();

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        closeDialog(dialog, "ok");
      }
    });
  });
}

type QuickPickItem<T = unknown> = { label: string; value: T; description?: string; detail?: string };

export async function showQuickPick<T>(
  items: QuickPickItem<T>[],
  options: { placeHolder?: string } = {},
): Promise<T | null> {
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

  return await new Promise<T | null>((resolve) => {
    dialog.addEventListener(
      "close",
      () => {
        const returnValue = String((dialog as any).returnValue ?? dialog.returnValue ?? "");
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

    let firstBtn: HTMLButtonElement | null = null;
    for (const [idx, item] of items.entries()) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "quick-pick__item";
      btn.dataset.testid = `quick-pick-item-${idx}`;
      if (!firstBtn) firstBtn = btn;

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

    showModal(dialog);
    try {
      firstBtn?.focus();
    } catch {
      // Best-effort focus.
    }
  });
}
