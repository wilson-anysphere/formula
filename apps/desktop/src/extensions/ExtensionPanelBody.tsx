import React from "react";

import type { ExtensionPanelBridge } from "./extensionPanelBridge.js";

const WEBVIEW_CSP = [
  // Disallow all network/resource loads by default. Extensions should bundle any assets
  // into the HTML they set (inline scripts/styles, data: URLs, etc). This avoids bypassing
  // the extension host permission model from inside the iframe.
  "default-src 'none'",
  "img-src data: blob:",
  "style-src 'unsafe-inline'",
  "script-src 'unsafe-inline'",
  "connect-src 'none'",
  "font-src data:",
  "base-uri 'none'",
  "form-action 'none'",
].join("; ");

function injectWebviewCsp(html: string): string {
  const cspMeta = `<meta http-equiv="Content-Security-Policy" content="${WEBVIEW_CSP}">`;
  const src = String(html ?? "");

  // If the extension already provides a `<head>`, inject our CSP as early as possible.
  const headMatch = /<head(\s[^>]*)?>/i.exec(src);
  if (headMatch && typeof headMatch.index === "number") {
    const insertAt = headMatch.index + headMatch[0].length;
    return `${src.slice(0, insertAt)}${cspMeta}${src.slice(insertAt)}`;
  }

  // Otherwise, inject a `<head>` right after the `<html>` element when present.
  const htmlMatch = /<html(\s[^>]*)?>/i.exec(src);
  if (htmlMatch && typeof htmlMatch.index === "number") {
    const insertAt = htmlMatch.index + htmlMatch[0].length;
    return `${src.slice(0, insertAt)}<head>${cspMeta}</head>${src.slice(insertAt)}`;
  }

  // Fall back to wrapping arbitrary markup in a full document.
  return `<!doctype html><html><head>${cspMeta}</head><body>${src}</body></html>`;
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
