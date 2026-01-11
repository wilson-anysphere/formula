export type CronSchedule = any;

export function parseCronExpression(expression: string): CronSchedule;

export function nextCronRun(schedule: CronSchedule, afterMs: number, timezone?: "local" | "utc"): number | null;
