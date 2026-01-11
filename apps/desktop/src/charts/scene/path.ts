import { formatNumber } from "./format.js";

export type PathCommand =
  | { kind: "M"; x: number; y: number }
  | { kind: "L"; x: number; y: number }
  | { kind: "AT"; x1: number; y1: number; x2: number; y2: number; radius: number }
  | { kind: "A"; cx: number; cy: number; r: number; startAngle: number; endAngle: number; ccw?: boolean }
  | { kind: "Z" };

export interface PathData {
  commands: PathCommand[];
}

export class PathBuilder {
  #commands: PathCommand[] = [];

  moveTo(x: number, y: number): this {
    this.#commands.push({ kind: "M", x, y });
    return this;
  }

  lineTo(x: number, y: number): this {
    this.#commands.push({ kind: "L", x, y });
    return this;
  }

  arcTo(x1: number, y1: number, x2: number, y2: number, radius: number): this {
    this.#commands.push({ kind: "AT", x1, y1, x2, y2, radius });
    return this;
  }

  arc(cx: number, cy: number, r: number, startAngle: number, endAngle: number, ccw = false): this {
    this.#commands.push({ kind: "A", cx, cy, r, startAngle, endAngle, ccw });
    return this;
  }

  closePath(): this {
    this.#commands.push({ kind: "Z" });
    return this;
  }

  build(): PathData {
    return { commands: [...this.#commands] };
  }
}

export function path(): PathBuilder {
  return new PathBuilder();
}

interface Point {
  x: number;
  y: number;
}

function nearlyEqual(a: number, b: number, eps = 1e-6): boolean {
  return Math.abs(a - b) <= eps;
}

function pointsNearlyEqual(a: Point, b: Point): boolean {
  return nearlyEqual(a.x, b.x) && nearlyEqual(a.y, b.y);
}

function computeArcToSegments(p0: Point, cmd: Extract<PathCommand, { kind: "AT" }>): { t1: Point; t2: Point; sweepFlag: 0 | 1 } | null {
  const { x1, y1, x2, y2 } = cmd;
  const r = cmd.radius;
  if (!Number.isFinite(r) || r <= 0) return null;

  const p1 = { x: x1, y: y1 };
  const p2 = { x: x2, y: y2 };
  if (pointsNearlyEqual(p0, p1) || pointsNearlyEqual(p1, p2)) return null;

  const v1 = { x: p0.x - p1.x, y: p0.y - p1.y };
  const v2 = { x: p2.x - p1.x, y: p2.y - p1.y };
  const len1 = Math.hypot(v1.x, v1.y);
  const len2 = Math.hypot(v2.x, v2.y);
  if (len1 === 0 || len2 === 0) return null;

  const u1 = { x: v1.x / len1, y: v1.y / len1 };
  const u2 = { x: v2.x / len2, y: v2.y / len2 };

  const dot = Math.max(-1, Math.min(1, u1.x * u2.x + u1.y * u2.y));
  const angle = Math.acos(dot);
  if (!Number.isFinite(angle) || angle <= 1e-6 || Math.abs(Math.PI - angle) <= 1e-6) return null;

  const t = r / Math.tan(angle / 2);
  const t1 = { x: p1.x + u1.x * t, y: p1.y + u1.y * t };
  const t2 = { x: p1.x + u2.x * t, y: p1.y + u2.y * t };

  const inDir = { x: p1.x - p0.x, y: p1.y - p0.y };
  const outDir = { x: p2.x - p1.x, y: p2.y - p1.y };
  const cross = inDir.x * outDir.y - inDir.y * outDir.x;
  const sweepFlag: 0 | 1 = cross > 0 ? 1 : 0;

  return { t1, t2, sweepFlag };
}

function polarPoint(cx: number, cy: number, r: number, angle: number): Point {
  return { x: cx + r * Math.cos(angle), y: cy + r * Math.sin(angle) };
}

function normalizeArcDelta(startAngle: number, endAngle: number, ccw: boolean): number {
  const tau = Math.PI * 2;
  let delta = endAngle - startAngle;
  if (!ccw) {
    if (delta <= 0) delta += tau;
  } else {
    if (delta >= 0) delta -= tau;
    delta = -delta;
  }
  if (!Number.isFinite(delta) || delta < 0) return 0;
  if (delta >= tau - 1e-6) return tau;
  return delta;
}

export function pathToSvgD(data: PathData): string {
  const parts: string[] = [];
  let current: Point | null = null;
  let subpathStart: Point | null = null;

  for (const cmd of data.commands) {
    switch (cmd.kind) {
      case "M": {
        const p = { x: cmd.x, y: cmd.y };
        parts.push(`M ${formatNumber(p.x)} ${formatNumber(p.y)}`);
        current = p;
        subpathStart = p;
        break;
      }
      case "L": {
        const p = { x: cmd.x, y: cmd.y };
        if (!current) {
          parts.push(`M ${formatNumber(p.x)} ${formatNumber(p.y)}`);
          current = p;
          subpathStart = p;
          break;
        }
        parts.push(`L ${formatNumber(p.x)} ${formatNumber(p.y)}`);
        current = p;
        break;
      }
      case "AT": {
        if (!current) {
          const p = { x: cmd.x1, y: cmd.y1 };
          parts.push(`M ${formatNumber(p.x)} ${formatNumber(p.y)}`);
          current = p;
          subpathStart = p;
          break;
        }

        const arc = computeArcToSegments(current, cmd);
        if (!arc) {
          const p = { x: cmd.x1, y: cmd.y1 };
          parts.push(`L ${formatNumber(p.x)} ${formatNumber(p.y)}`);
          current = p;
          break;
        }

        if (!pointsNearlyEqual(current, arc.t1)) {
          parts.push(`L ${formatNumber(arc.t1.x)} ${formatNumber(arc.t1.y)}`);
        }
        parts.push(
          `A ${formatNumber(cmd.radius)} ${formatNumber(cmd.radius)} 0 0 ${arc.sweepFlag} ${formatNumber(arc.t2.x)} ${formatNumber(arc.t2.y)}`
        );
        current = arc.t2;
        break;
      }
      case "A": {
        const start = polarPoint(cmd.cx, cmd.cy, cmd.r, cmd.startAngle);
        const end = polarPoint(cmd.cx, cmd.cy, cmd.r, cmd.endAngle);
        if (!current) {
          parts.push(`M ${formatNumber(start.x)} ${formatNumber(start.y)}`);
          current = start;
          subpathStart = start;
        } else if (!pointsNearlyEqual(current, start)) {
          parts.push(`L ${formatNumber(start.x)} ${formatNumber(start.y)}`);
          current = start;
        }

        const ccw = cmd.ccw ?? false;
        const delta = normalizeArcDelta(cmd.startAngle, cmd.endAngle, ccw);
        const sweepFlag = ccw ? 0 : 1;

        if (delta === 0) break;

        if (nearlyEqual(delta, Math.PI * 2)) {
          const midAngle = cmd.startAngle + (ccw ? -Math.PI : Math.PI);
          const mid = polarPoint(cmd.cx, cmd.cy, cmd.r, midAngle);
          parts.push(
            `A ${formatNumber(cmd.r)} ${formatNumber(cmd.r)} 0 0 ${sweepFlag} ${formatNumber(mid.x)} ${formatNumber(mid.y)}`
          );
          parts.push(
            `A ${formatNumber(cmd.r)} ${formatNumber(cmd.r)} 0 0 ${sweepFlag} ${formatNumber(start.x)} ${formatNumber(start.y)}`
          );
          current = start;
          break;
        }

        const largeArcFlag: 0 | 1 = delta > Math.PI ? 1 : 0;
        parts.push(
          `A ${formatNumber(cmd.r)} ${formatNumber(cmd.r)} 0 ${largeArcFlag} ${sweepFlag} ${formatNumber(end.x)} ${formatNumber(end.y)}`
        );
        current = end;
        break;
      }
      case "Z": {
        parts.push("Z");
        current = subpathStart;
        break;
      }
    }
  }

  return parts.join(" ");
}

export function applyPathToCanvas(ctx: CanvasRenderingContext2D, data: PathData): void {
  let hasPoint = false;

  for (const cmd of data.commands) {
    switch (cmd.kind) {
      case "M":
        ctx.moveTo(cmd.x, cmd.y);
        hasPoint = true;
        break;
      case "L":
        if (!hasPoint) {
          ctx.moveTo(cmd.x, cmd.y);
          hasPoint = true;
          break;
        }
        ctx.lineTo(cmd.x, cmd.y);
        break;
      case "AT":
        if (!hasPoint) {
          ctx.moveTo(cmd.x1, cmd.y1);
          hasPoint = true;
          break;
        }
        ctx.arcTo(cmd.x1, cmd.y1, cmd.x2, cmd.y2, cmd.radius);
        break;
      case "A":
        ctx.arc(cmd.cx, cmd.cy, cmd.r, cmd.startAngle, cmd.endAngle, cmd.ccw ?? false);
        hasPoint = true;
        break;
      case "Z":
        ctx.closePath();
        break;
    }
  }
}

