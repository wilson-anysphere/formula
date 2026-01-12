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
  "script-src 'unsafe-inline'",
  "connect-src 'none'",
  "worker-src 'none'",
  "child-src 'none'",
  "frame-src 'none'",
  "font-src data:",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

export function injectWebviewCsp(html: string): string {
  const cspMeta = `<meta http-equiv="Content-Security-Policy" content="${WEBVIEW_CSP}">`;
  const hardenTauriGlobalsScript = `<script>
 (() => {
  "use strict";
  const keys = ["__TAURI__", "__TAURI_IPC__", "__TAURI_INTERNALS__", "__TAURI_METADATA__"];
  let tauriGlobalsPresent = false;

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
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
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

      try {
        // If we couldn't fully delete the global, attempt to lock it down to undefined so it can't be
        // re-populated later in the page lifecycle.
        Object.defineProperty(window, key, {
          value: undefined,
          writable: false,
          configurable: false,
        });
        continue;
      } catch {
        // Ignore.
      }

      try {
        // If deletion fails, fall back to overwriting.
        window[key] = undefined;
      } catch {
        // Ignore.
      }

      try {
        // Best-effort: after overwriting, try to lock down the property to prevent later reinjection.
        Object.defineProperty(window, key, {
          value: undefined,
          writable: false,
          configurable: false,
        });
      } catch {
        // Ignore.
      }
    }
  };

  // Run immediately, then schedule a few additional scrubs in case globals are injected after
  // the initial script executes (best-effort defense-in-depth).
  scrubTauriGlobals();
  try {
    Promise.resolve().then(scrubTauriGlobals);
  } catch {
    // Ignore.
  }
  try {
    setTimeout(scrubTauriGlobals, 0);
    setTimeout(scrubTauriGlobals, 50);
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
</script>`;
  const injectedHeadContent = `${cspMeta}${hardenTauriGlobalsScript}`;
  const src = String(html ?? "");

  // If the markup doesn't include an `<html>`, wrap it so the iframe always runs in standards mode
  // and the CSP meta tag is guaranteed to be parsed as a head directive.
  if (!/<html(\s[^>]*)?>/i.test(src)) {
    return `<!doctype html><html><head>${injectedHeadContent}</head><body>${src}</body></html>`;
  }

  // Ensure the CSP applies before any extension-provided markup is parsed (including malformed
  // HTML that places tags before `<html>`/`<head>`). We do this by injecting immediately after the
  // document doctype when present, otherwise by prefixing a doctype and injecting right after it.
  const withDoctype = /^\s*<!doctype\b/i.test(src) ? src : `<!doctype html>${src}`;
  const doctypeMatch = /^\s*<!doctype[^>]*>/i.exec(withDoctype);
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
      referrerPolicy="no-referrer"
      src={src ?? undefined}
      style={{ width: "100%", height: "100%", border: "0", background: "transparent" }}
    />
  );
}
