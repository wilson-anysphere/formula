import React from "react";

import type { ExtensionPanelBridge } from "./extensionPanelBridge.js";

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

  const documentHtml = html && html.trim().length > 0 ? html : "<!doctype html><html><body></body></html>";

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
