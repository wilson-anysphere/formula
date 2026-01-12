export type ClipboardContent = {
  text?: string;
  html?: string;
  rtf?: string;
  pngBase64?: string;
};

export type ClipboardWritePayload = {
  text?: string;
  html?: string;
  rtf?: string;
  pngBase64?: string;
};

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

export async function readClipboard(): Promise<ClipboardContent> {
  const invoke = getTauriInvoke();
  const payload = await invoke("clipboard_read");
  return (payload as ClipboardContent) ?? {};
}

export async function writeClipboard(payload: ClipboardWritePayload): Promise<void> {
  const invoke = getTauriInvoke();
  await invoke("clipboard_write", { payload });
}

