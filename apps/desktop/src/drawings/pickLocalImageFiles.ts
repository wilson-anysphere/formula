import { MAX_INSERT_IMAGE_BYTES } from "./insertImageLimits.js";
import { pickImagesFromTauriDialog, readBinaryFile } from "./pickImagesFromTauriDialog.js";

export interface PickLocalImageFilesOptions {
  /**
   * Allow selecting multiple images.
   *
   * Defaults to true since the Ribbon UI supports inserting multiple pictures at once.
   */
  multiple?: boolean;
}

/**
 * Open a native file picker (via `<input type="file">`) and resolve with selected image files.
 *
 * - Non-blocking: returns a Promise and does not block the main thread.
 * - Best-effort cancel handling: resolves `[]` when the user dismisses the picker.
 * - Cleans up the temporary input element and listeners.
 */
export function pickLocalImageFiles(options: PickLocalImageFilesOptions = {}): Promise<File[]> {
  const tauriDialogOpenAvailable =
    typeof (globalThis as any).__TAURI__?.dialog?.open === "function" ||
    typeof (globalThis as any).__TAURI__?.plugin?.dialog?.open === "function" ||
    typeof (globalThis as any).__TAURI__?.plugins?.dialog?.open === "function";
  const tauriInvokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";

  // Desktop/Tauri: prefer the native dialog + backend file reads so we can work with
  // filesystem paths directly (avoids `<input type=file>` sandbox quirks).
  if (tauriDialogOpenAvailable && tauriInvokeAvailable) {
    return pickLocalImageFilesViaTauriDialog(options);
  }

  if (typeof document === "undefined" || !document.body) return Promise.resolve([]);

  return new Promise((resolve, reject) => {
    const input = document.createElement("input");
    input.type = "file";
    input.multiple = options.multiple ?? true;
    // Prefer the broad accept string, but include explicit extensions so platforms
    // with weaker MIME inference still show common image formats.
    input.accept = "image/*,.png,.jpg,.jpeg,.gif,.bmp,.webp,.svg";

    // Keep it out of view but still "clickable".
    input.style.position = "fixed";
    input.style.left = "-9999px";
    input.style.top = "0";

    document.body.appendChild(input);

    let settled = false;
    const cleanup = () => {
      input.remove();
      window.removeEventListener("focus", onWindowFocus, true);
    };

    const finish = (files: File[]) => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(files);
    };

    const onChange = () => {
      finish(input.files ? Array.from(input.files) : []);
    };

    // Some browsers (Chrome) fire a `cancel` event for `<input type="file">`. It's not
    // universally supported, so we also use a focus-based fallback below.
    const onCancel = () => {
      finish([]);
    };

    // Best-effort cancel detection:
    // Opening the file picker typically blurs the window; when the picker is dismissed
    // the window regains focus. If no `change` event fired, treat it as cancel.
    const onWindowFocus = () => {
      // Defer a tick to allow the `change` event to win the race when a file is selected.
      setTimeout(() => {
        if (settled) return;
        finish(input.files ? Array.from(input.files) : []);
      }, 0);
    };

    input.addEventListener("change", onChange, { once: true });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    input.addEventListener("cancel" as any, onCancel as any, { once: true });
    window.addEventListener("focus", onWindowFocus, true);

    try {
      input.click();
    } catch (err) {
      cleanup();
      reject(err);
    }
  });
}

async function pickLocalImageFilesViaTauriDialog(options: PickLocalImageFilesOptions): Promise<File[]> {
  const multiple = options.multiple ?? true;
  const paths = await pickImagesFromTauriDialog();
  const selected = multiple ? paths : paths.slice(0, 1);
  if (selected.length === 0) return [];

  const FileCtor = (globalThis as any).File as typeof File | undefined;
  if (typeof FileCtor !== "function") {
    throw new Error("File API not available in this environment");
  }

  const out: File[] = [];
  for (const path of selected) {
    const name = basename(path);
    try {
      const bytes = await readBinaryFile(path);
      const mimeType = guessImageMimeType(name);
      out.push(new FileCtor([bytes], name, { type: mimeType }));
    } catch (err) {
      const message = String((err as any)?.message ?? err);
      // Oversized images should be skipped (with user feedback handled by the insertion layer).
      // Return a lightweight placeholder with a `size` above the cap so callers can filter + toast
      // without allocating the full byte payload.
      if (message.toLowerCase().includes("file is too large")) {
        out.push({ name, type: guessImageMimeType(name), size: MAX_INSERT_IMAGE_BYTES + 1 } as any as File);
        continue;
      }
      throw err;
    }
  }
  return out;
}

function basename(path: string): string {
  const raw = String(path ?? "");
  const parts = raw.split(/[/\\]/);
  return parts[parts.length - 1] ?? raw;
}

function guessImageMimeType(name: string): string {
  const ext = String(name ?? "")
    .split(".")
    .pop()
    ?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    default:
      return "application/octet-stream";
  }
}
