const formula = require("@formula/extension-api");

async function activate(context) {
  let panel = null;
  let panelListener = null;

  async function ensurePanel() {
    if (panel) return panel;

    panel = await formula.ui.createPanel("sampleHello.panel", {
      title: "Sample Hello Panel",
      position: "right"
    });

    panelListener = panel.webview.onDidReceiveMessage(async (message) => {
      if (message && message.type === "ping") {
        await panel.webview.postMessage({ type: "pong" });
      }
    });
    context.subscriptions.push(panelListener);

    return panel;
  }

  async function renderPanel() {
    const p = await ensurePanel();
    await p.webview.setHtml(`<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8" />
    <title>Sample Hello Panel</title>
  </head>
  <body>
    <h1>Sample Hello Panel</h1>
    <p>This panel is rendered from an extension.</p>
  </body>
</html>`);
    return p.id;
  }

  const doubleFn = await formula.functions.register("SAMPLEHELLO_DOUBLE", {
    description: "Doubles the input value",
    parameters: [{ name: "value", type: "number", description: "Value to double" }],
    result: { type: "number" },
    handler: (value) => {
      if (typeof value !== "number" || !Number.isFinite(value)) return null;
      return value * 2;
    }
  });

  const sumCmd = await formula.commands.registerCommand("sampleHello.sumSelection", async () => {
    const selection = await formula.cells.getSelection();
    const values = selection.values ?? [];

    let sum = 0;
    for (const row of values) {
      for (const val of row) {
        if (typeof val === "number" && Number.isFinite(val)) sum += val;
      }
    }

    await formula.cells.setCell(selection.endRow + 1, selection.startCol, sum);
    await formula.ui.showMessage(`Sum: ${sum}`, "info");
    return sum;
  });

  const panelCmd = await formula.commands.registerCommand("sampleHello.openPanel", async () => {
    return renderPanel();
  });

  const viewActivated = formula.events.onViewActivated(async ({ viewId }) => {
    if (viewId === "sampleHello.panel") {
      await renderPanel();
    }
  });

  const fetchCmd = await formula.commands.registerCommand("sampleHello.fetchText", async (url) => {
    const response = await fetch(String(url));
    const text = await response.text();
    await formula.ui.showMessage(`Fetched: ${text}`, "info");
    return text;
  });

  const copySumCmd = await formula.commands.registerCommand(
    "sampleHello.copySumToClipboard",
    async () => {
      const selection = await formula.cells.getSelection();
      const values = selection.values ?? [];

      let sum = 0;
      for (const row of values) {
        for (const val of row) {
          if (typeof val === "number" && Number.isFinite(val)) sum += val;
        }
      }

      await formula.clipboard.writeText(String(sum));
      await formula.ui.showMessage(`Copied sum ${sum} to clipboard`, "info");
      return sum;
    }
  );

  context.subscriptions.push(doubleFn, sumCmd, panelCmd, fetchCmd, copySumCmd, viewActivated);
}

module.exports = {
  activate
};
