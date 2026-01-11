export type LocaleDirection = "ltr" | "rtl";

export function availableLocales(): string[];
export function getLocale(): string;
export function getDirection(locale?: string): LocaleDirection;
export function setLocale(locale: string): void;
export function t(key: string): string;
export function tWithVars(key: string, vars?: Record<string, unknown>): string;

