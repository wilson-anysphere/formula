import type { DrawingTransform } from "./types";

const ROTATION_UNITS_PER_DEGREE = 60_000;

export function normalizeRotationDeg(rotationDeg: number): number {
  if (!Number.isFinite(rotationDeg)) return 0;
  // DrawingML accepts rotations outside [0, 360). Normalize to [-180, 180).
  let normalized = rotationDeg % 360;
  if (normalized >= 180) normalized -= 360;
  if (normalized < -180) normalized += 360;
  // Avoid -0 which can be annoying in equality checks / snapshots.
  return Object.is(normalized, -0) ? 0 : normalized;
}

function parseBoolAttr(raw: string | undefined): boolean {
  if (!raw) return false;
  const value = raw.trim().toLowerCase();
  if (value === "" || value === "0" || value === "false" || value === "f") return false;
  return true;
}

function parseNumberAttr(raw: string | undefined): number | null {
  if (!raw) return null;
  const value = Number(raw.trim());
  return Number.isFinite(value) ? value : null;
}

function parseAttributes(source: string): Record<string, string> {
  // NOTE: This is intentionally a tiny attribute parser (no entity decoding, etc).
  // The DrawingML snippets we preserve are machine-generated and typically use
  // straightforward `key="value"` attributes.
  const attrs: Record<string, string> = {};
  // Support both single and double quotes to match real-world OOXML payloads
  // (Excel can emit either, and some round-trip fixtures use single quotes).
  const attrRe = /([A-Za-z_][\w:.-]*)\s*=\s*(['"])([^'"]*)\2/g;
  for (;;) {
    const match = attrRe.exec(source);
    if (!match) break;
    const [, key, _quote, value] = match;
    attrs[key] = value;
  }
  return attrs;
}

/**
 * Best-effort extraction of DrawingML transform metadata from an object `raw_xml`
 * snippet (e.g. `<xdr:sp>…</xdr:sp>`).
 */
export function parseDrawingTransformFromRawXml(rawXml: string): DrawingTransform | null {
  if (typeof rawXml !== "string" || rawXml.length === 0) return null;

  // Prefer `a:xfrm` (shape properties). Fall back to `xdr:xfrm` (graphicFrame).
  const xfrmMatch =
    rawXml.match(/<a:xfrm\b([^>]*)>/) ??
    rawXml.match(/<xdr:xfrm\b([^>]*)>/) ??
    rawXml.match(/<xfrm\b([^>]*)>/);

  if (!xfrmMatch) return null;
  const attrs = parseAttributes(xfrmMatch[1] ?? "");

  const rotRaw = parseNumberAttr(attrs.rot);
  const rotationDeg =
    rotRaw == null ? 0 : normalizeRotationDeg(rotRaw / ROTATION_UNITS_PER_DEGREE);

  const flipH = parseBoolAttr(attrs.flipH);
  const flipV = parseBoolAttr(attrs.flipV);

  return { rotationDeg, flipH, flipV };
}

export function degToRad(deg: number): number {
  return (deg * Math.PI) / 180;
}

export function rotateVector(x: number, y: number, radians: number): { x: number; y: number } {
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return { x: x * cos - y * sin, y: x * sin + y * cos };
}

export function inverseTransformVector(
  dx: number,
  dy: number,
  transform: DrawingTransform,
): { x: number; y: number } {
  // Inverse of: rotate(theta) then scale(flip).
  const rotated = rotateVector(dx, dy, -degToRad(transform.rotationDeg));
  return {
    x: transform.flipH ? -rotated.x : rotated.x,
    y: transform.flipV ? -rotated.y : rotated.y,
  };
}

export function applyTransformVector(
  dx: number,
  dy: number,
  transform: DrawingTransform,
): { x: number; y: number } {
  // Forward transform: scale(flip) then rotate(theta).
  const sx = transform.flipH ? -1 : 1;
  const sy = transform.flipV ? -1 : 1;
  const scaled = { x: dx * sx, y: dy * sy };
  return rotateVector(scaled.x, scaled.y, degToRad(transform.rotationDeg));
}
