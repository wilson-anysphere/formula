export type Rect = { r0: number; c0: number; r1: number; c1: number };

export function rectSize(rect: Rect): number;
export function rectIntersectionArea(a: Rect, b: Rect): number;
export function cellToA1(row0: number, col0: number): string;
export function rectToA1(rect: Rect): string;
