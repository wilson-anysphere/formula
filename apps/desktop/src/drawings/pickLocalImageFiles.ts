import { FileTooLargeError, pickImagesFromTauriDialog, readBinaryFile } from "./pickImagesFromTauriDialog.js";
import { getTauriDialogOpenOrNull, getTauriInvokeOrNull } from "../tauri/api";

export interface PickLocalImageFilesOptions {
  /**
   * Allow selecting multiple images.
   *
   * Defaults to true since the Ribbon UI supports inserting multiple pictures at once.
   */
  multiple?: boolean;
}

/**
 * Pick local image files for insertion.
 *
 * - Desktop/Tauri: uses the native file dialog (via `getTauriDialogOpenOrNull()` from `tauri/api`)
 *   and reads bytes via hardened backend invoke commands.
 * - Web: falls back to a hidden `<input type="file">`.
 *
 * - Non-blocking: returns a Promise and does not block the main thread.
 * - Best-effort cancel handling: resolves `[]` when the user dismisses the picker.
 * - Cleans up the temporary input element and listeners.
 */
export function pickLocalImageFiles(options: PickLocalImageFilesOptions = {}): Promise<File[]> {
  const tauriDialogOpenAvailable = getTauriDialogOpenOrNull() != null;
  const tauriInvokeAvailable = getTauriInvokeOrNull() != null;

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
    let focusTimer: number | null = null;
    const cleanup = () => {
      input.remove();
      window.removeEventListener("focus", onWindowFocus, true);
      if (focusTimer != null) {
        clearTimeout(focusTimer);
        focusTimer = null;
      }
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
      // Browsers usually blur the window while the native file picker is open and then
      // re-focus when it's dismissed. If the user cancelled, no `change` event fires.
      //
      // In automation (Playwright) the focus transition can happen before the harness
      // calls `fileChooser.setFiles(...)`. Use a small delay so a real selection's
      // `change` event can win the race.
      if (focusTimer != null) clearTimeout(focusTimer);
      focusTimer = window.setTimeout(() => {
        focusTimer = null;
        if (settled) return;
        finish(input.files ? Array.from(input.files) : []);
      }, 250);
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
  const paths = await pickImagesFromTauriDialog({ multiple });
  const selected = multiple ? paths : paths.slice(0, 1);
  if (selected.length === 0) return [];

  const FileCtor = (globalThis as any).File as typeof File | undefined;
  if (typeof FileCtor !== "function") {
    throw new Error("File API not available in this environment");
  }

  const out: File[] = [];
  for (const path of selected) {
    const name = basename(path);
    const mimeType = guessImageMimeType(name);
    try {
      const bytes = await readBinaryFile(path);
      out.push(new FileCtor([bytes], name, { type: mimeType }));
    } catch (err) {
      // If the file is too large, return an oversized placeholder File so callers can
      // surface an appropriate UI error without forcing us to load the whole payload.
      if (err instanceof FileTooLargeError) {
        try {
          const placeholder = new FileCtor([new Uint8Array(0)], name, { type: mimeType });
          let patched = false;
          try {
            // `File.size` is a getter on the prototype chain; define an own property so the
            // placeholder compares as oversized without allocating large buffers.
            Object.defineProperty(placeholder, "size", { value: err.fileSize });
            patched = placeholder.size === err.fileSize;
          } catch {
            patched = false;
          }
          if (patched) {
            out.push(placeholder);
          } else {
            out.push({ name, type: mimeType, size: err.maxSize + 1 } as any as File);
          }
        } catch {
          // Some environments may not allow constructing/shimming a `File` instance. Fall back
          // to a lightweight File-like object with an oversized `size`.
          out.push({ name, type: mimeType, size: err.maxSize + 1 } as any as File);
        }
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
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}
