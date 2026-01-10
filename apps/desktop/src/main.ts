import { SpreadsheetApp } from "./app/spreadsheetApp";

const gridRoot = document.getElementById("grid");
if (!gridRoot) {
  throw new Error("Missing #grid container");
}

const activeCell = document.querySelector<HTMLElement>('[data-testid="active-cell"]');
const selectionRange = document.querySelector<HTMLElement>('[data-testid="selection-range"]');
const activeValue = document.querySelector<HTMLElement>('[data-testid="active-value"]');
const openComments = document.querySelector<HTMLButtonElement>('[data-testid="open-comments-panel"]');
if (!activeCell || !selectionRange || !activeValue) {
  throw new Error("Missing status bar elements");
}
if (!openComments) {
  throw new Error("Missing comments panel toggle button");
}

const app = new SpreadsheetApp(gridRoot, { activeCell, selectionRange, activeValue });
app.focus();
openComments.addEventListener("click", () => app.toggleCommentsPanel());

// Expose a small API for Playwright assertions.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(window as any).__formulaApp = app;
