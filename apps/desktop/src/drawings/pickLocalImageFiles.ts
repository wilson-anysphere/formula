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

