import type { SheetNameResolver } from "../../sheet/sheetNameResolver";

export class StructuralConflictUiController {
  [key: string]: any;
  constructor(opts: {
    container: HTMLElement;
    monitor: { resolveConflict: (id: string, resolution: any) => boolean };
    sheetNameResolver?: SheetNameResolver | null | undefined;
    onNavigateToCell?: ((cellRef: { sheetId: string; row: number; col: number }) => void) | undefined;
    resolveUserLabel?: ((userId: string) => string) | undefined;
  });

  destroy(): void;
  addConflict(conflict: any): void;
  render(): void;
}
