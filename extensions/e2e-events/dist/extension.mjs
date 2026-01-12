import * as formula from "@formula/extension-api";

function safeSet(key, value) {
  try {
    // Fire-and-forget: event handlers are invoked synchronously by the runtime.
    void formula.storage.set(String(key), value).catch(() => {
      // ignore
    });
  } catch {
    // ignore
  }
}

export async function activate(context) {
  // Reset last-known payloads so e2e assertions don't read stale data if the same
  // browser context is reused.
  safeSet("selectionChanged", null);
  safeSet("cellChanged", null);
  safeSet("sheetActivated", null);
  safeSet("workbookOpened", null);
  safeSet("beforeSave", null);
  safeSet("viewActivated", null);

  context.subscriptions.push(
    formula.events.onSelectionChanged((e) => {
      safeSet("selectionChanged", e);
    })
  );

  context.subscriptions.push(
    formula.events.onCellChanged((e) => {
      safeSet("cellChanged", e);
    })
  );

  context.subscriptions.push(
    formula.events.onSheetActivated((e) => {
      safeSet("sheetActivated", e);
    })
  );

  context.subscriptions.push(
    formula.events.onWorkbookOpened((e) => {
      safeSet("workbookOpened", e);
    })
  );

  context.subscriptions.push(
    formula.events.onBeforeSave((e) => {
      safeSet("beforeSave", e);
    })
  );

  context.subscriptions.push(
    formula.events.onViewActivated((e) => {
      safeSet("viewActivated", e);
    })
  );
}

export default { activate };
