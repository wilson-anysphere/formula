/**
 * Lightweight helpers for classifying DrawingML XML payloads.
 *
 * The desktop renderer intentionally avoids a full XML parse for drawing objects
 * because these helpers run in the WebView bundle (no Node APIs) and can be
 * called frequently while rendering.
 */

/**
 * Returns true when the raw DrawingML payload contains a `graphicFrame` element.
 *
 * In SpreadsheetDrawing, `xdr:graphicFrame` is used for a variety of embedded
 * objects (charts, SmartArt/diagrams, etc).
 */
export function isGraphicFrame(rawXml: string | null | undefined): boolean {
  if (!rawXml) return false;
  // Match `<xdr:graphicFrame ...>` as well as other namespace prefixes.
  return /<\s*(?:[A-Za-z0-9_-]+:)?graphicFrame\b/.test(rawXml);
}

/**
 * Best-effort classification for diagram-based graphic frames (SmartArt).
 */
export function isSmartArtGraphicFrame(rawXml: string | null | undefined): boolean {
  if (!rawXml) return false;
  if (!isGraphicFrame(rawXml)) return false;
  // SmartArt uses `a:graphicData uri=".../diagram"` and/or `dgm:*` elements.
  return rawXml.includes("drawingml/2006/diagram") || /<\s*dgm:/.test(rawXml);
}

/**
 * Placeholder label for unsupported graphic frames.
 */
export function graphicFramePlaceholderLabel(rawXml: string | null | undefined): "SmartArt" | "GraphicFrame" | null {
  if (!rawXml) return null;
  if (!isGraphicFrame(rawXml)) return null;
  return isSmartArtGraphicFrame(rawXml) ? "SmartArt" : "GraphicFrame";
}

