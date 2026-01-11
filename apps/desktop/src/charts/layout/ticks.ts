import { normalizeZero, round } from "./geometry";

export function formatTickValue(value: number, formatCode?: string | null): string {
  if (!Number.isFinite(value)) return "";

  const code = (formatCode ?? "").split(";")[0]?.trim();
  if (!code || code.toLowerCase() === "general") {
    const rounded = Math.round(value * 100) / 100;
    if (Number.isInteger(rounded)) return String(rounded);
    return String(rounded.toFixed(2)).replace(/\.?0+$/, "");
  }

  const isPercent = code.includes("%");
  const wantsThousands = code.includes(",");
  const raw = isPercent ? value * 100 : value;

  let decimals = 0;
  const dotIndex = code.indexOf(".");
  if (dotIndex >= 0) {
    const afterDot = code.slice(dotIndex + 1);
    const match = afterDot.match(/^[0#]+/);
    decimals = match ? match[0].length : 0;
  }

  let out = raw.toFixed(decimals);
  if (wantsThousands) out = addThousandsSeparators(out);
  if (isPercent) out += "%";
  return out;
}

function addThousandsSeparators(num: string): string {
  const [intPart, frac] = num.split(".");
  const sign = intPart.startsWith("-") ? "-" : "";
  const digits = sign ? intPart.slice(1) : intPart;
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return frac ? `${sign}${grouped}.${frac}` : `${sign}${grouped}`;
}

function niceStep(roughStep: number): number {
  if (!Number.isFinite(roughStep) || roughStep <= 0) return 1;
  const pow10 = 10 ** Math.floor(Math.log10(roughStep));
  const frac = roughStep / pow10;
  const niceFrac = frac <= 1 ? 1 : frac <= 2 ? 2 : frac <= 5 ? 5 : 10;
  return niceFrac * pow10;
}

function floorToStep(value: number, step: number): number {
  return Math.floor(value / step) * step;
}

function ceilToStep(value: number, step: number): number {
  return Math.ceil(value / step) * step;
}

export interface TickGenerationResult {
  domain: [number, number];
  ticks: number[];
}

/**
 * Generate a deterministic "Excel-like" set of major tick values.
 *
 * v1 heuristic:
 * - Target 5-7 ticks.
 * - Use 1/2/5*10^k steps (d3-ish).
 * - For fully explicit axis bounds, fall back to evenly-spaced ticks so the
 *   provided min/max remain the domain extents.
 */
export function generateLinearTicks(args: {
  domain: [number, number];
  minExplicit: boolean;
  maxExplicit: boolean;
  tickCountMin?: number;
  tickCountMax?: number;
}): TickGenerationResult {
  const tickCountMin = args.tickCountMin ?? 5;
  const tickCountMax = args.tickCountMax ?? 7;
  let [min, max] = args.domain;

  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    min = 0;
    max = 1;
  }

  if (min === max) {
    const bumped = min === 0 ? 1 : Math.abs(min * 0.1);
    min -= bumped;
    max += bumped;
  }

  if (min > max) [min, max] = [max, min];

  // When both bounds are explicit, don't "nice" the domain.
  if (args.minExplicit && args.maxExplicit) {
    const ticks = linearTicks(min, max, clampInt(6, tickCountMin, tickCountMax));
    return { domain: [normalizeZero(min), normalizeZero(max)], ticks };
  }

  const span = max - min;
  const target = 6;

  /** @type {TickGenerationResult | null} */
  let best = null;
  let bestScore = Number.POSITIVE_INFINITY;

  for (const desired of [5, 6, 7]) {
    const step = niceStep(span / (desired - 1));
    const start = args.minExplicit ? min : floorToStep(min, step);
    const end = args.maxExplicit ? max : ceilToStep(max, step);

    const ticks = rangeTicks(start, end, step);
    const count = ticks.length;
    const inRange = count >= tickCountMin && count <= tickCountMax;
    const score = (inRange ? 0 : 100) + Math.abs(count - target);
    if (score < bestScore) {
      bestScore = score;
      best = { domain: [normalizeZero(start), normalizeZero(end)], ticks };
    }
  }

  return best ?? { domain: [normalizeZero(min), normalizeZero(max)], ticks: linearTicks(min, max, target) };
}

function clampInt(n: number, min: number, max: number): number {
  if (!Number.isFinite(n)) return min;
  return Math.max(min, Math.min(max, Math.trunc(n)));
}

function linearTicks(min: number, max: number, count: number): number[] {
  if (count <= 1) return [normalizeZero(round(min, 12))];
  const step = (max - min) / (count - 1);
  return Array.from({ length: count }, (_, i) => normalizeZero(round(min + i * step, 12)));
}

function rangeTicks(start: number, end: number, step: number): number[] {
  const ticks: number[] = [];
  if (!Number.isFinite(start) || !Number.isFinite(end) || !Number.isFinite(step) || step <= 0) return ticks;
  // Include end tick even with floating point accumulation noise.
  const n = Math.floor((end - start) / step + 0.5) + 1;
  for (let i = 0; i < n; i += 1) {
    ticks.push(normalizeZero(round(start + i * step, 12)));
  }
  // Ensure last tick is exactly the (rounded) end when end is auto-derived.
  if (ticks.length) {
    const last = ticks[ticks.length - 1];
    const roundedEnd = normalizeZero(round(end, 12));
    if (Math.abs(last - roundedEnd) > Math.abs(step) / 1e6) {
      ticks.push(roundedEnd);
    }
  }
  return ticks;
}

