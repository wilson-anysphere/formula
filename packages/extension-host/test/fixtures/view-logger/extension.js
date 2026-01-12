// Minimal extension for extension-host integration tests.
// It activates on startup and logs view activation events via ui.showMessage.
"use strict";

const formula = require("@formula/extension-api");

async function activate(context) {
  context.subscriptions.push(
    formula.events.onViewActivated(({ viewId }) => {
      void formula.ui.showMessage(
        `[view-logger] viewActivated:${typeof viewId}:${String(viewId)}`,
      );
    })
  );
}

module.exports = { activate };
