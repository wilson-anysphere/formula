import * as React from "react";

import { showToast } from "../extensions/ui.js";

type Props = {
  pdfBytes: Uint8Array;
  filename: string;
  autoPrint?: boolean;
  onDownload: () => void;
  onClose: () => void;
};

export function PrintPreviewDialog({ pdfBytes, filename, autoPrint = false, onDownload, onClose }: Props) {
  const iframeRef = React.useRef<HTMLIFrameElement | null>(null);
  const printButtonRef = React.useRef<HTMLButtonElement | null>(null);

  const [viewerLoaded, setViewerLoaded] = React.useState(false);
  const autoPrintedRef = React.useRef(false);

  const pdfUrl = React.useMemo(() => {
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    return URL.createObjectURL(blob);
  }, [pdfBytes]);

  React.useEffect(() => {
    return () => {
      try {
        URL.revokeObjectURL(pdfUrl);
      } catch {
        // ignore
      }
    };
  }, [pdfUrl]);

  React.useEffect(() => {
    // Ensure keyboard users land on the primary action.
    requestAnimationFrame(() => printButtonRef.current?.focus());
  }, []);

  const tryPrint = React.useCallback((): boolean => {
    try {
      const win = iframeRef.current?.contentWindow;
      if (!win) return false;

      // Accessing `print` can throw in some WebView / cross-origin iframe situations.
      let fn: any = null;
      try {
        fn = (win as any)?.print;
      } catch {
        return false;
      }
      if (typeof fn !== "function") return false;

      try {
        (win as any).focus?.();
      } catch {
        // ignore focus issues (WebView restrictions)
      }

      try {
        fn.call(win);
        return true;
      } catch {
        return false;
      }
    } catch {
      return false;
    }
  }, []);

  const handlePrint = React.useCallback(() => {
    const ok = tryPrint();
    if (ok) return;

    // Best-effort fallback: download the PDF and ask the user to print from their viewer.
    try {
      onDownload();
    } catch {
      // ignore
    }
    showToast(`Couldn't open the print dialog automatically. Downloaded ${filename} instead; please print it using your PDF viewer.`);
  }, [filename, onDownload, tryPrint]);

  React.useEffect(() => {
    if (!autoPrint) return;
    if (!viewerLoaded) return;
    if (autoPrintedRef.current) return;
    autoPrintedRef.current = true;

    // Defer so the iframe has a chance to finish initializing its PDF viewer.
    window.setTimeout(() => {
      try {
        handlePrint();
      } catch {
        // ignore
      }
    }, 0);
  }, [autoPrint, handlePrint, viewerLoaded]);

  return (
    <div className="print-preview-dialog__root">
      <div className="print-preview-dialog__header">
        <div className="print-preview-dialog__title">Print Preview</div>
        <div className="dialog__controls print-preview-dialog__controls">
          <button ref={printButtonRef} type="button" onClick={handlePrint}>
            Print
          </button>
          <button type="button" onClick={onDownload}>
            Download PDF
          </button>
          <button type="button" onClick={onClose}>
            Close
          </button>
        </div>
      </div>
      <iframe
        ref={iframeRef}
        className="print-preview-dialog__viewer"
        src={pdfUrl}
        title="Print Preview"
        onLoad={() => setViewerLoaded(true)}
      />
    </div>
  );
}
