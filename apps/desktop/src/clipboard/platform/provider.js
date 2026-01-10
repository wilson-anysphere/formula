/**
 * Platform clipboard integration.
 *
 * Desktop (Tauri): prefers the Tauri clipboard API when available.
 * Web: uses the browser Clipboard API (permission-gated).
 *
 * Note: Clipboard access generally requires a user gesture in browsers. Callers
 * should invoke these APIs from explicit copy/paste handlers.
 *
 * @typedef {import("../types.js").ClipboardContent} ClipboardContent
 * @typedef {import("../types.js").ClipboardWritePayload} ClipboardWritePayload
 *
 * @typedef {{
 *   read: () => Promise<ClipboardContent>,
 *   write: (payload: ClipboardWritePayload) => Promise<void>
 * }} ClipboardProvider
 */

function hasTauri() {
  return Boolean(globalThis.__TAURI__);
}

/**
 * @returns {Promise<ClipboardProvider>}
 */
export async function createClipboardProvider() {
  if (hasTauri()) return createTauriClipboardProvider();
  return createWebClipboardProvider();
}

/**
 * @returns {ClipboardProvider}
 */
function createTauriClipboardProvider() {
  const tauriClipboard = globalThis.__TAURI__?.clipboard;

  return {
    async read() {
      if (tauriClipboard?.readText) {
        const text = await tauriClipboard.readText();
        return { text: text ?? undefined };
      }
      // Fall back to browser text read when the plugin is unavailable.
      return createWebClipboardProvider().read();
    },
    async write(payload) {
      if (tauriClipboard?.writeText) {
        await tauriClipboard.writeText(payload.text);
      } else {
        await createWebClipboardProvider().write({ text: payload.text });
      }

      // Best-effort HTML write via ClipboardItem when available (WebView-dependent).
      if (payload.html && typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await navigator.clipboard.write([item]);
        } catch {
          // Ignore; some platforms deny HTML clipboard writes.
        }
      }
    },
  };
}

/**
 * @returns {ClipboardProvider}
 */
function createWebClipboardProvider() {
  return {
    async read() {
      // Prefer rich read if available.
      if (navigator.clipboard?.read) {
        try {
          const items = await navigator.clipboard.read();
          for (const item of items) {
            const htmlType = item.types.find((t) => t === "text/html");
            const textType = item.types.find((t) => t === "text/plain");

            const html =
              htmlType &&
              (await item.getType(htmlType).then((b) => b.text()).catch(() => undefined));
            const text =
              textType &&
              (await item.getType(textType).then((b) => b.text()).catch(() => undefined));

            if (html || text) return { html, text };
          }
        } catch {
          // Permission denied or unsupported – fall back to plain text below.
        }
      }

      let text;
      try {
        text = navigator.clipboard?.readText ? await navigator.clipboard.readText() : undefined;
      } catch {
        text = undefined;
      }
      return { text };
    },
    async write(payload) {
      // Prefer writing both formats when possible.
      if (payload.html && typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await navigator.clipboard.write([item]);
          return;
        } catch {
          // Fall back to plain text.
        }
      }

      if (navigator.clipboard?.writeText) {
        try {
          await navigator.clipboard.writeText(payload.text);
        } catch {
          // Ignore; clipboard write requires user gesture/permissions in browsers.
        }
      }
    },
  };
}

