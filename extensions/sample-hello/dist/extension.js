const formula = require("@formula/extension-api");

async function activate(context) {
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
    const panel = await formula.ui.createPanel("sampleHello.panel", {
      title: "Sample Hello Panel",
      position: "right"
    });

    await panel.webview.setHtml(`<!DOCTYPE html>
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

    return panel.id;
  });

  context.subscriptions.push(sumCmd, panelCmd);
}

module.exports = {
  activate
};

