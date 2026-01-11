import type { Rect } from "./types";

export function clampNonNegative(n: number): number {
  if (!Number.isFinite(n)) return 0;
  return n < 0 ? 0 : n;
}

export function normalizeZero(n: number): number {
  return Object.is(n, -0) ? 0 : n;
}

export function round(n: number, decimals = 6): number {
  if (!Number.isFinite(n)) return n;
  const m = 10 ** decimals;
  return normalizeZero(Math.round(n * m) / m);
}

export function insetRect(rect: Rect, inset: number): Rect {
  return {
    x: rect.x + inset,
    y: rect.y + inset,
    width: Math.max(0, rect.width - inset * 2),
    height: Math.max(0, rect.height - inset * 2),
  };
}

export function rectRight(rect: Rect): number {
  return rect.x + rect.width;
}

export function rectBottom(rect: Rect): number {
  return rect.y + rect.height;
}

