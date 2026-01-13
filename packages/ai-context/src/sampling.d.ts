export type SeededRng = () => number;

export interface RandomSamplingOptions {
  seed?: number;
  rng?: SeededRng;
}

export interface StratifiedSamplingOptions<T> extends RandomSamplingOptions {
  getStratum: (row: T) => string;
}

export interface SystematicSamplingOptions extends RandomSamplingOptions {
  /**
   * Fractional offset within the sampling interval.
   *
   * - `0` means start at the beginning of the first interval
   * - `0.5` means start halfway through the first interval
   *
   * When omitted, the offset is derived deterministically from `seed` / `rng`.
   */
  offset?: number;
}

export function createSeededRng(seed: number): SeededRng;
export function randomSampleIndices(total: number, sampleSize: number, rng: SeededRng): number[];
export function randomSampleRows<T>(rows: T[], sampleSize: number, options?: RandomSamplingOptions): T[];
export function headSampleRows<T>(rows: T[], sampleSize: number): T[];
export function tailSampleRows<T>(rows: T[], sampleSize: number): T[];
export function systematicSampleRows<T>(
  rows: T[],
  sampleSize: number,
  options?: SystematicSamplingOptions,
): T[];
export function stratifiedSampleRows<T>(rows: T[], sampleSize: number, options: StratifiedSamplingOptions<T>): T[];
