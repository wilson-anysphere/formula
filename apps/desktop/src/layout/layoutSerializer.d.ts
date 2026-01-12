export type PanelRegistryLike =
  | Record<string, unknown>
  | {
      has?: (panelId: string) => boolean;
      hasPanel?: (panelId: string) => boolean;
    };

export function serializeLayout(
  layout: unknown,
  options?: { panelRegistry?: PanelRegistryLike | null; primarySheetId?: string | null },
): string;

export function deserializeLayout(
  json: string,
  options?: { panelRegistry?: PanelRegistryLike | null; primarySheetId?: string | null },
): any;
