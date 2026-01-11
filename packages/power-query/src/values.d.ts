export const MS_PER_DAY: number;

export function parseIsoLikeToUtcDate(input: string): Date | null;

export function hasUtcTimeComponent(value: Date): boolean;

export class PqDecimal {
  value: string;

  constructor(value: string);

  toString(): string;
}

export function isPqDecimal(value: unknown): value is PqDecimal;

export class PqTime {
  milliseconds: number;

  constructor(milliseconds: number);
}

export function isPqTime(value: unknown): value is PqTime;

export class PqDuration {
  milliseconds: number;

  constructor(milliseconds: number);
}

export function isPqDuration(value: unknown): value is PqDuration;

export class PqDateTimeZone {
  constructor(input: Date | string);

  toDate(): Date;
}

export function isPqDateTimeZone(value: unknown): value is PqDateTimeZone;
