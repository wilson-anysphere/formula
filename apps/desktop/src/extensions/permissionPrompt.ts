import * as nativeDialogs from "../tauri/nativeDialogs.js";

type NetworkRequestInfo = {
  host?: string;
  url?: string;
  mode?: string;
};

export type ExtensionPermissionPromptRequest = {
  extensionId: string;
  displayName?: string;
  permissions: string[];
  apiKey?: string;
  request?: {
    apiKey?: string;
    permissions?: string[];
    network?: NetworkRequestInfo;
  };
};

export type ExtensionPermissionPrompt = (req: ExtensionPermissionPromptRequest) => Promise<boolean>;

let permissionPromptTitleId = 0;
function nextPermissionPromptTitleId(): string {
  permissionPromptTitleId += 1;
  return `extension-permission-prompt-title-${permissionPromptTitleId}`;
}

function showModal(dialog: HTMLDialogElement): void {
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
    return;
  }
  // jsdom doesn't implement showModal(). Best-effort fallback so unit tests can
  // drive the DOM without a full dialog polyfill.
  dialog.setAttribute("open", "true");
}

function closeDialog(dialog: HTMLDialogElement, returnValue: string): void {
  if (typeof dialog.close === "function") {
    dialog.close(returnValue);
    return;
  }
  // jsdom doesn't implement close()/returnValue. Emulate the close contract so
  // our prompt logic and tests can rely on the close event.
  (dialog as any).returnValue = returnValue;
  dialog.removeAttribute("open");
  dialog.dispatchEvent(new Event("close"));
}

function formatPermissionLabel(permission: string, req: ExtensionPermissionPromptRequest): string {
  if (permission === "network") {
    const host = req.request?.network?.host;
    const url = req.request?.network?.url;
    if (host) return `network (${host})`;
    if (url) return `network (${url})`;
  }

  return permission;
}

async function promptOnce(req: ExtensionPermissionPromptRequest): Promise<boolean> {
  const permissions = Array.isArray(req.permissions) ? req.permissions.map(String) : [];
  if (permissions.length === 0) return true;

  if (typeof document === "undefined" || !document.body) {
    const ext = req.displayName ? `${req.displayName} (${req.extensionId})` : req.extensionId;
    const list = permissions.map((p) => formatPermissionLabel(p, req)).join(", ");
    return nativeDialogs.confirm(`Allow ${ext} to use: ${list}?`, { title: "Extension Permission Request" });
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog";
  dialog.dataset.testid = "extension-permission-prompt";
  dialog.dataset.keybindingBarrier = "true";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = "Extension Permission Request";
  title.id = nextPermissionPromptTitleId();
  dialog.setAttribute("aria-labelledby", title.id);

  const extensionMeta = document.createElement("div");
  extensionMeta.className = "permission-prompt__meta";

  const extensionName = document.createElement("div");
  extensionName.className = "permission-prompt__name";
  extensionName.textContent = req.displayName ? String(req.displayName) : String(req.extensionId);

  const extensionId = document.createElement("div");
  extensionId.className = "permission-prompt__id";
  extensionId.textContent = String(req.extensionId);

  extensionMeta.appendChild(extensionName);
  if (req.displayName && req.displayName !== req.extensionId) {
    extensionMeta.appendChild(extensionId);
  }

  const detail = document.createElement("div");
  detail.textContent = "This extension is requesting access to:";
  detail.className = "permission-prompt__detail";

  const list = document.createElement("ul");
  list.className = "permission-prompt__list";

  for (const perm of permissions) {
    const li = document.createElement("li");
    li.textContent = formatPermissionLabel(perm, req);
    li.dataset.testid = `extension-permission-${perm}`;
    list.appendChild(li);
  }

  const note = document.createElement("div");
  note.className = "permission-prompt__note";
  note.textContent = "Your choice will be saved for this extension.";

  const controls = document.createElement("div");
  controls.className = "dialog__controls permission-prompt__controls";

  const denyBtn = document.createElement("button");
  denyBtn.type = "button";
  denyBtn.textContent = "Deny";
  denyBtn.dataset.testid = "extension-permission-deny";

  const allowBtn = document.createElement("button");
  allowBtn.type = "button";
  allowBtn.textContent = "Allow";
  allowBtn.dataset.testid = "extension-permission-allow";

  controls.appendChild(denyBtn);
  controls.appendChild(allowBtn);

  dialog.appendChild(title);
  dialog.appendChild(extensionMeta);
  dialog.appendChild(detail);
  dialog.appendChild(list);
  dialog.appendChild(note);
  dialog.appendChild(controls);

  document.body.appendChild(dialog);

  const cleanup = () => {
    dialog.remove();
  };

  return await new Promise<boolean>((resolve) => {
    dialog.addEventListener(
      "close",
      () => {
        const returnValue = String((dialog as any).returnValue ?? dialog.returnValue ?? "");
        cleanup();
        resolve(returnValue === "allow");
      },
      { once: true },
    );

    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      closeDialog(dialog, "deny");
    });

    denyBtn.addEventListener("click", () => closeDialog(dialog, "deny"));
    allowBtn.addEventListener("click", () => closeDialog(dialog, "allow"));

    showModal(dialog);
    allowBtn.focus();
  });
}

/**
 * Creates a permission prompt implementation suitable for wiring into the
 * BrowserExtensionHost. Calls are automatically serialized so multiple
 * permission requests don't stack multiple modal dialogs.
 */
export function createDesktopPermissionPrompt(): ExtensionPermissionPrompt {
  let running = false;
  const pending: Array<{
    req: ExtensionPermissionPromptRequest;
    resolve: (result: boolean) => void;
    reject: (err: unknown) => void;
  }> = [];

  const drain = async () => {
    if (running) return;
    running = true;
    try {
      while (pending.length > 0) {
        const next = pending.shift();
        if (!next) continue;
        try {
          // eslint-disable-next-line no-await-in-loop
          const result = await promptOnce(next.req);
          next.resolve(result);
        } catch (err) {
          next.reject(err);
        }
      }
    } finally {
      running = false;
    }
  };

  return (req) =>
    new Promise<boolean>((resolve, reject) => {
      pending.push({ req, resolve, reject });
      void drain();
    });
}
