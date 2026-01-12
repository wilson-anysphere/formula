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
 * @param {ClipboardContent} content
 */
function hasAnyContent(content) {
  return Boolean(
    content &&
      (content.text !== undefined ||
        content.html !== undefined ||
        content.rtf !== undefined ||
        content.imagePng !== undefined)
  );
}

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function encodeBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  if (typeof globalThis.btoa !== "function") {
    throw new Error("Base64 encoding unavailable: missing Buffer and btoa");
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return globalThis.btoa(binary);
}

/**
 * @param {string} base64
 * @returns {Uint8Array}
 */
function decodeBase64(base64) {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(base64, "base64"));
  }

  if (typeof globalThis.atob !== "function") {
    throw new Error("Base64 decoding unavailable: missing Buffer and atob");
  }

  const binary = globalThis.atob(base64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

/**
 * Merge `incoming` into `base` without clobbering defined keys on `base`.
 *
 * @param {ClipboardContent} base
 * @param {ClipboardContent} incoming
 * @returns {ClipboardContent}
 */
function mergeClipboardContent(base, incoming) {
  /** @type {ClipboardContent} */
  const out = { ...base };

  if (out.text === undefined && incoming.text !== undefined) out.text = incoming.text;
  if (out.html === undefined && incoming.html !== undefined) out.html = incoming.html;
  if (out.rtf === undefined && incoming.rtf !== undefined) out.rtf = incoming.rtf;
  if (out.imagePng === undefined && incoming.imagePng !== undefined) out.imagePng = incoming.imagePng;

  return out;
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
  const tauriCore = globalThis.__TAURI__?.core;

  return {
    async read() {
      // Prefer rich reads via the WebView Clipboard API when available so we can
      // ingest HTML tables + formats from external spreadsheets.
      const web = await createWebClipboardProvider({
        // Prefer native Tauri `readText` for plain text when available; use the Web
        // `readText` fallback only when we don't have a native alternative.
        fallbackToReadText: !tauriClipboard?.readText,
      }).read();
      let merged = web;
      if (tauriCore?.invoke) {
        try {
          const data = await tauriCore.invoke("read_clipboard");
          if (data && typeof data === "object") {
            /** @type {any} */
            const anyData = data;
            /** @type {ClipboardContent} */
            const invokeContent = {
              text: typeof anyData.text === "string" ? anyData.text : undefined,
              html: typeof anyData.html === "string" ? anyData.html : undefined,
              rtf: typeof anyData.rtf === "string" ? anyData.rtf : undefined,
              imagePng:
                typeof anyData.image_png_base64 === "string"
                  ? (() => {
                      try {
                        return decodeBase64(anyData.image_png_base64);
                      } catch {
                        return undefined;
                      }
                    })()
                  : undefined,
            };
            merged = mergeClipboardContent(web, invokeContent);
          }
        } catch {
          // Ignore; bridge command may not exist yet.
        }
      }

      if (hasAnyContent(merged)) return merged;

      if (tauriClipboard?.readText) {
        const text = await tauriClipboard.readText();
        return { text: text ?? undefined };
      }

      return merged;
    },
    async write(payload) {
      const hasNonText =
        payload.html !== undefined || payload.rtf !== undefined || payload.imagePng !== undefined;

      // Prefer the native Tauri clipboard API for plain text when available.
      if (!hasNonText && tauriClipboard?.writeText) {
        try {
          await tauriClipboard.writeText(payload.text);
        } catch {
          // Fall back to Web clipboard below.
          await createWebClipboardProvider().write({ text: payload.text });
        }
        return;
      }

      // Best-effort: ensure text is written even if rich formats fail.
      if (tauriClipboard?.writeText) {
        try {
          await tauriClipboard.writeText(payload.text);
        } catch {
          // Ignore; continue with rich write attempts.
        }
      }

      if (tauriCore?.invoke && (hasNonText || !tauriClipboard?.writeText)) {
        try {
          await tauriCore.invoke("write_clipboard", {
            text: payload.text,
            html: payload.html,
            rtf: payload.rtf,
            image_png_base64: payload.imagePng ? encodeBase64(payload.imagePng) : undefined,
          });
          return;
        } catch {
          // Ignore; bridge command may not exist yet.
        }
      }

      // No rich Tauri bridge available; fall back to the Web Clipboard API (best effort).
      await createWebClipboardProvider().write(payload);
    },
  };
}

/**
 * @returns {ClipboardProvider}
 */
function createWebClipboardProvider(options = {}) {
  const { fallbackToReadText = true } = options;

  return {
    async read() {
      // Prefer rich read if available.
      const clipboard = globalThis.navigator?.clipboard;
      if (clipboard?.read) {
        try {
          const items = await clipboard.read();
          for (const item of items) {
            const types = Array.isArray(item.types) ? item.types : [];
            const htmlType = types.find((t) => t === "text/html");
            const textType = types.find((t) => t === "text/plain");
            const rtfType = types.find((t) => t === "text/rtf");
            const imageType = types.find((t) => t === "image/png");

            if (!htmlType && !textType && !rtfType && !imageType) continue;

            const html =
              htmlType &&
              (await item.getType(htmlType).then((b) => b.text()).catch(() => undefined));
            const text =
              textType &&
              (await item.getType(textType).then((b) => b.text()).catch(() => undefined));
            const rtf =
              rtfType && (await item.getType(rtfType).then((b) => b.text()).catch(() => undefined));
            const imagePng =
              imageType &&
              (await item
                .getType(imageType)
                .then((b) => b.arrayBuffer())
                .then((ab) => new Uint8Array(ab))
                .catch(() => undefined));

            if (html !== undefined || text !== undefined || rtf !== undefined || imagePng !== undefined) {
              return { html, text, rtf, imagePng };
            }
          }
        } catch {
          // Permission denied or unsupported – fall back to plain text below.
        }
      }

      if (!fallbackToReadText) return {};

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

      // Prefer writing rich formats when possible.
      if (typeof ClipboardItem !== "undefined" && clipboard?.write) {
        try {
          /** @type {Record<string, Blob>} */
          const data = {
            "text/plain": new Blob([payload.text], { type: "text/plain" }),
          };

          if (payload.html !== undefined) {
            data["text/html"] = new Blob([payload.html], { type: "text/html" });
          }
          if (payload.rtf !== undefined) {
            data["text/rtf"] = new Blob([payload.rtf], { type: "text/rtf" });
          }
          if (payload.imagePng !== undefined) {
            data["image/png"] = new Blob([payload.imagePng], { type: "image/png" });
          }

          const item = new ClipboardItem(data);
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
