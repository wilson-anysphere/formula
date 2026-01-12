// Global declarations for the desktop Playwright/Vitest harness.
//
// These globals are assigned in `apps/desktop/src/main.ts` so e2e tests can
// introspect/drive the app without going through UI affordances.
//
// NOTE: Keep these as `unknown` to avoid coupling tests to internal types.
export {};

declare global {
  interface Window {
    __formulaApp?: unknown;
    __formulaCommandRegistry?: unknown;
    __formulaExtensionHostManager?: unknown;
    __formulaExtensionHost?: unknown;
    __formulaSecondaryGrid?: unknown;
    __workbookSheetStore?: unknown;
  }
}
