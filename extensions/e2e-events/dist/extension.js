"use strict";

const formula = require("@formula/extension-api");

function safeSet(key, value) {
  try {
    Promise.resolve(formula.storage.set(String(key), value)).catch(() => {
      // ignore
    });
  } catch {
    // ignore
  }
}

async function activate(context) {
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

module.exports = {
  activate,
};

