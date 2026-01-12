// Global declarations for the desktop e2e harness.
//
// These globals are assigned in `apps/desktop/src/main.ts` so Playwright/Vitest tests can
// introspect/drive the app without going through UI affordances.
//
// NOTE: Keep these as `unknown` to avoid coupling tests (or app code) to internal types.
export {};

declare global {
  interface Window {
    __formulaApp?: unknown;
    __formulaCommandRegistry?: unknown;
    __formulaExtensionHostManager?: unknown;
    __formulaExtensionHost?: unknown;
    __formulaSecondaryGrid?: unknown;
    __layoutController?: unknown;
    __workbookSheetStore?: unknown;
  }
}
