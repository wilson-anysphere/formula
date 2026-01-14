import React from "react";

import type { ExtensionPanelBridge } from "./extensionPanelBridge.js";

const WEBVIEW_CSP = [
  // Disallow all network/resource loads by default. Extensions should bundle any assets
  // into the HTML they set (inline scripts/styles, data: URLs, etc). This avoids bypassing
  // the extension host permission model from inside the iframe.
  "default-src 'none'",
  "object-src 'none'",
  "img-src data: blob:",
  "style-src 'unsafe-inline'",
  // Desktop/Tauri CSP disallows inline scripts. Allow only self-contained module sources.
  "script-src blob: data:",
  "connect-src 'none'",
  "worker-src 'none'",
  "child-src 'none'",
  "frame-src 'none'",
  "font-src data:",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

const HARDEN_TAURI_GLOBALS_SOURCE = `(() => {
  "use strict";
  const keys = ["__TAURI__", "__TAURI_IPC__", "__TAURI_INTERNALS__", "__TAURI_METADATA__", "__TAURI_INVOKE__"];
  let tauriGlobalsPresent = false;
  const areTauriGlobalsScrubbed = () => {
    for (const key of keys) {
      try {
        if (typeof window[key] !== "undefined") return false;
      } catch {
        return false;
      }
    }
    return true;
  };

  const lockDownKey = (key) => {
    let enumerable = false;
    let isAccessor = false;
    let configurable = true;
    try {
      const desc = Object.getOwnPropertyDescriptor(window, key);
      if (desc) {
        enumerable = Boolean(desc.enumerable);
        configurable = Boolean(desc.configurable);
        isAccessor = typeof desc.get === "function" || typeof desc.set === "function";
      }
    } catch {
      // Ignore.
    }

    try {
      if (isAccessor) {
        // Can't redefine an accessor if it's non-configurable.
        if (!configurable) return false;
        Object.defineProperty(window, key, {
          get: () => undefined,
          set: () => {},
          enumerable,
          configurable: false,
        });
        return true;
      }

      Object.defineProperty(window, key, {
        value: undefined,
        writable: false,
        enumerable,
        configurable: false,
      });
      return true;
    } catch {
      return false;
    }
  };

  const scrubTauriGlobals = () => {
    for (const key of keys) {
      let hasKey = false;
      try {
        hasKey = key in window;
      } catch {
        // Ignore.
      }
      if (!hasKey) continue;

      tauriGlobalsPresent = true;
      let deleted = false;
      try {
        // Some environments (including strict mode) throw if the property is non-configurable.
        deleted = delete window[key];
      } catch {
        // Ignore.
      }

      if (deleted) {
        let stillPresent = false;
        try {
          stillPresent = key in window;
        } catch {
          stillPresent = true;
        }
        if (!stillPresent) continue;
      }

      // If we couldn't fully delete the global, attempt to lock it down to undefined so it can't be
      // re-populated later in the page lifecycle.
      if (lockDownKey(key)) continue;

      try {
        // If deletion fails, fall back to overwriting.
        window[key] = undefined;
      } catch {
        // Ignore.
      }

      // Best-effort: after overwriting, try to lock down the property to prevent later reinjection.
      lockDownKey(key);
    }

  };

  // Run immediately, then schedule a few additional scrubs in case globals are injected after
  // the initial script executes (best-effort defense-in-depth).
  scrubTauriGlobals();
  try {
    Promise.resolve()
      .then(scrubTauriGlobals)
      .catch(() => {
        // Best-effort: avoid unhandled rejections if the scrub callback throws.
      });
  } catch {
    // Ignore.
  }
  try {
    setTimeout(scrubTauriGlobals, 0);
    setTimeout(scrubTauriGlobals, 50);
    setTimeout(scrubTauriGlobals, 250);
    setTimeout(scrubTauriGlobals, 1000);
    document.addEventListener("DOMContentLoaded", scrubTauriGlobals, { once: true });
    window.addEventListener("load", scrubTauriGlobals, { once: true });
  } catch {
    // Ignore.
  }
  try {
    // Marker used by desktop e2e tests to verify the hardening script ran in the iframe.
    const marker = {};
    try {
      Object.defineProperty(marker, "tauriGlobalsPresent", {
        enumerable: true,
        get: () => tauriGlobalsPresent,
      });
    } catch {
      marker.tauriGlobalsPresent = tauriGlobalsPresent;
    }
    try {
      Object.defineProperty(marker, "tauriGlobalsScrubbed", {
        enumerable: true,
        get: () => areTauriGlobalsScrubbed(),
      });
    } catch {
      marker.tauriGlobalsScrubbed = areTauriGlobalsScrubbed();
    }
    try {
      Object.freeze(marker);
    } catch {
      // Ignore.
    }
    try {
      Object.defineProperty(window, "__formulaWebviewSandbox", {
        value: marker,
        writable: false,
        configurable: false,
      });
    } catch {
      window.__formulaWebviewSandbox = marker;
    }
  } catch {
    // Ignore.
  }
})();
`;

// Avoid inline scripts: Tauri CSP blocks 'unsafe-inline', and blob: documents inherit the
// parent document CSP. Use a `data:` URL script so it can execute under the default policy.
const HARDEN_TAURI_GLOBALS_SCRIPT_URL = `data:text/javascript;charset=utf-8,${encodeURIComponent(
  HARDEN_TAURI_GLOBALS_SOURCE,
)}`;
const HARDEN_TAURI_GLOBALS_SCRIPT = `<script src="${HARDEN_TAURI_GLOBALS_SCRIPT_URL}"></script>`;
const WEBVIEW_CSP_META = `<meta http-equiv="Content-Security-Policy" content="${WEBVIEW_CSP}">`;
const WEBVIEW_HEAD_INJECTION = `${WEBVIEW_CSP_META}${HARDEN_TAURI_GLOBALS_SCRIPT}`;

export function injectWebviewCsp(html: string): string {
  const injectedHeadContent = WEBVIEW_HEAD_INJECTION;
  const src = String(html ?? "");

  // Ensure the CSP applies before any extension-provided markup is parsed (including malformed
  // HTML that places tags before `<html>`/`<head>`). We do this by injecting immediately after the
  // document doctype when present, otherwise by prefixing a doctype and injecting right after it.
  const hasDoctype = /^[\uFEFF\s]*(?:<!--[\s\S]*?-->[\uFEFF\s]*)*<!doctype\b/i.test(src);
  const withDoctype = hasDoctype ? src : `<!doctype html>${src}`;
  const doctypeMatch = /^[\uFEFF\s]*(?:<!--[\s\S]*?-->[\uFEFF\s]*)*<!doctype[^>]*>/i.exec(withDoctype);
  if (doctypeMatch) {
    const insertAt = doctypeMatch.index + doctypeMatch[0].length;
    return `${withDoctype.slice(0, insertAt)}${injectedHeadContent}${withDoctype.slice(insertAt)}`;
  }

  return `${injectedHeadContent}${withDoctype}`;
}

export function ExtensionPanelBody({ panelId, bridge }: { panelId: string; bridge: ExtensionPanelBridge }) {
  const [html, setHtml] = React.useState(() => bridge.getPanelHtml(panelId));
  const [src, setSrc] = React.useState<string | null>(null);
  const iframeRef = React.useRef<HTMLIFrameElement | null>(null);

  React.useEffect(() => {
    const unsubscribe = bridge.subscribe(panelId, () => {
      setHtml(bridge.getPanelHtml(panelId));
    });
    // Ensure we don't miss a setHtml update that arrives between first render and effect subscription.
    setHtml(bridge.getPanelHtml(panelId));
    return unsubscribe;
  }, [bridge, panelId]);

  const setIframe = React.useCallback(
    (node: HTMLIFrameElement | null) => {
      if (iframeRef.current) bridge.disconnect(panelId, iframeRef.current);
      iframeRef.current = node;
      if (node) bridge.connect(panelId, node);
    },
    [bridge, panelId],
  );

  const documentHtml = injectWebviewCsp(
    html && html.trim().length > 0 ? html : "<!doctype html><html><body></body></html>",
  );

  React.useEffect(() => {
    const blob = new Blob([documentHtml], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    setSrc(url);
    return () => {
      URL.revokeObjectURL(url);
    };
  }, [documentHtml]);

  return (
    <iframe
      ref={setIframe}
      title={panelId}
      data-testid={`extension-webview-${panelId}`}
      sandbox="allow-scripts"
      allow="clipboard-read 'none'; clipboard-write 'none'; camera 'none'; microphone 'none'; geolocation 'none'"
      referrerPolicy="no-referrer"
      src={src ?? undefined}
      className="extension-webview"
    />
  );
}
