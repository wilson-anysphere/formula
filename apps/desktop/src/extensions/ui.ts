export type MessageType = "info" | "warning" | "error";

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
  dialog.className = "dialog";
  dialog.dataset.testid = "input-box";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = options.prompt ?? "Input";

  const input = document.createElement("input");
  input.type = "text";
  input.value = options.value ?? "";
  if (options.placeHolder) input.placeholder = options.placeHolder;
  input.style.width = "100%";
  input.style.padding = "10px 12px";
  input.style.borderRadius = "10px";
  input.style.border = "1px solid var(--border)";
  input.style.background = "var(--bg-primary)";
  input.style.color = "var(--text-primary)";
  input.dataset.testid = "input-box-field";

  const controls = document.createElement("div");
  controls.style.display = "flex";
  controls.style.justifyContent = "flex-end";
  controls.style.gap = "8px";
  controls.style.marginTop = "12px";

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
        const returnValue = dialog.returnValue;
        cleanup();
        resolve(returnValue === "ok" ? input.value : null);
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      dialog.close("cancel");
    });

    cancelBtn.addEventListener("click", () => dialog.close("cancel"));
    okBtn.addEventListener("click", () => dialog.close("ok"));

    dialog.showModal();
    input.focus();

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        dialog.close("ok");
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
  dialog.className = "dialog";
  dialog.dataset.testid = "quick-pick";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = options.placeHolder ?? "Select an item";

  const list = document.createElement("div");
  list.style.display = "flex";
  list.style.flexDirection = "column";
  list.style.gap = "6px";
  list.style.marginTop = "10px";

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
        const returnValue = dialog.returnValue;
        cleanup();
        if (!returnValue) resolve(null);
        else resolve(items[Number(returnValue)]?.value ?? null);
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      dialog.close("");
    });

    for (const [idx, item] of items.entries()) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.style.textAlign = "left";
      btn.style.padding = "10px 12px";
      btn.style.borderRadius = "10px";
      btn.style.border = "1px solid var(--border)";
      btn.style.background = "var(--bg-primary)";
      btn.style.color = "var(--text-primary)";
      btn.dataset.testid = `quick-pick-item-${idx}`;

      const label = document.createElement("div");
      label.textContent = item.label;
      label.style.fontWeight = "600";
      btn.appendChild(label);

      const secondary = item.description ?? item.detail ?? null;
      if (secondary) {
        const sub = document.createElement("div");
        sub.textContent = secondary;
        sub.style.color = "var(--text-secondary)";
        sub.style.fontSize = "12px";
        btn.appendChild(sub);
      }

      btn.addEventListener("click", () => {
        dialog.close(String(idx));
      });

      list.appendChild(btn);
    }

    dialog.showModal();
  });
}

