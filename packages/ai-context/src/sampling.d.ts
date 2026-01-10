export type SeededRng = () => number;

export interface RandomSamplingOptions {
  seed?: number;
  rng?: SeededRng;
}

export interface StratifiedSamplingOptions<T> extends RandomSamplingOptions {
  getStratum: (row: T) => string;
}

export function createSeededRng(seed: number): SeededRng;
export function randomSampleIndices(total: number, sampleSize: number, rng: SeededRng): number[];
export function randomSampleRows<T>(rows: T[], sampleSize: number, options?: RandomSamplingOptions): T[];
export function stratifiedSampleRows<T>(rows: T[], sampleSize: number, options: StratifiedSamplingOptions<T>): T[];

