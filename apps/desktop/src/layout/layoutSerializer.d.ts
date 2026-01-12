export function serializeLayout(layout: unknown, options?: { panelRegistry?: Record<string, unknown>; primarySheetId?: string | null }): string;

export function deserializeLayout(
  json: string,
  options?: { panelRegistry?: Record<string, unknown>; primarySheetId?: string | null },
): any;

