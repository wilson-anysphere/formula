import { ScriptRuntime, Workbook } from "@formula/scripting/web";

// Expose the scripting runtime for Playwright e2e tests.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).__formulaScripting = { ScriptRuntime, Workbook };

