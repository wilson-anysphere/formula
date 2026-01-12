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
  const tauriInvoke = globalThis.__TAURI__?.core?.invoke;
  const tauriClipboard = globalThis.__TAURI__?.clipboard;

  return {
    async read() {
      // 1) Prefer rich reads via the native clipboard command when available (Tauri IPC).
      if (typeof tauriInvoke === "function") {
        try {
          const result = await tauriInvoke("clipboard_read");
          if (result && typeof result === "object") {
            /** @type {any} */
            const r = result;
            const html = typeof r.html === "string" ? r.html : undefined;
            const text = typeof r.text === "string" ? r.text : undefined;
            const rtf = typeof r.rtf === "string" ? r.rtf : undefined;
            const pngBase64 =
              typeof r.pngBase64 === "string"
                ? r.pngBase64
                : typeof r.png_base64 === "string"
                  ? r.png_base64
                  : undefined;

            if (typeof html === "string" || typeof text === "string") {
              return { html, text, rtf, pngBase64 };
            }
          }
        } catch {
          // Ignore; command may not exist on older builds.
        }
      }

      // 2) Fall back to rich reads via the WebView Clipboard API when available so we can
      // ingest HTML tables + formats from external spreadsheets.
      const web = await createWebClipboardProvider().read();
      if (typeof web.html === "string" || typeof web.text === "string") return web;

      // 3) Final fallback: legacy Tauri plain text clipboard API.
      if (tauriClipboard?.readText) {
        try {
          const text = await tauriClipboard.readText();
          return { text: text ?? undefined };
        } catch {
          // Ignore.
        }
      }

      return web;
    },

    async write(payload) {
      // 1) Prefer rich writes via the native clipboard command when available (Tauri IPC).
      let wrote = false;
      if (typeof tauriInvoke === "function") {
        try {
          await tauriInvoke("clipboard_write", {
            text: payload.text,
            html: payload.html,
            rtf: payload.rtf,
            pngBase64: payload.pngBase64,
          });
          wrote = true;
        } catch {
          wrote = false;
        }
      }

      if (!wrote) {
        if (tauriClipboard?.writeText) {
          try {
            await tauriClipboard.writeText(payload.text);
          } catch {
            await createWebClipboardProvider().write({ text: payload.text });
          }
        } else {
          await createWebClipboardProvider().write({ text: payload.text });
        }
      }

      // 2) Secondary path: Best-effort HTML write via ClipboardItem when available (WebView-dependent).
      const clipboard = globalThis.navigator?.clipboard;
      if (payload.html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await clipboard.write([item]);
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
      const clipboard = globalThis.navigator?.clipboard;

      // Prefer rich read if available.
      if (clipboard?.read) {
        try {
          const items = await clipboard.read();
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
        text = clipboard?.readText ? await clipboard.readText() : undefined;
      } catch {
        text = undefined;
      }
      return { text };
    },

    async write(payload) {
      const clipboard = globalThis.navigator?.clipboard;

      // Prefer writing both formats when possible.
      if (payload.html && typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          const item = new ClipboardItem({
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
            "text/html": new Blob([payload.html], { type: "text/html" }),
          });
          await clipboard.write([item]);
          return;
        } catch {
          // Fall back to plain text.
        }
      }

      if (clipboard?.writeText) {
        try {
          await clipboard.writeText(payload.text);
        } catch {
          // Ignore; clipboard write requires user gesture/permissions in browsers.
        }
      }
    },
  };
}
