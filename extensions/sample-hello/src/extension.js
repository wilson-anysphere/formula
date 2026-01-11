const formula = require("@formula/extension-api");
const { sumValues } = require("./util");

const PANEL_ID = "sampleHello.panel";
const PANEL_TITLE = "Sample Hello Panel";
const CONNECTOR_ID = "sampleHello.connector";

/** @type {import("@formula/extension-api").Panel | null} */
let panel = null;
/** @type {Promise<import("@formula/extension-api").Panel> | null} */
let panelPromise = null;
/** @type {import("@formula/extension-api").Disposable | null} */
let panelMessageDisposable = null;

function panelHtml() {
  return `<!doctype html>
<html>
  <body>
    <h1>Sample Hello Panel</h1>
  </body>
</html>`;
}

/**
async function getSelectionSum() {
  const selection = await formula.cells.getSelection();
  return sumValues(selection?.values);
}

/**
 * @param {import("@formula/extension-api").ExtensionContext} context
 */
async function ensurePanel(context) {
  const html = panelHtml();

  if (panel) {
    try {
      await panel.webview.setHtml(html);
      return panel;
    } catch (error) {
      panel = null;
      if (panelMessageDisposable) {
        try {
          panelMessageDisposable.dispose();
        } catch {
          // ignore
        }
        panelMessageDisposable = null;
      }
    }
  }

  if (panelPromise) return panelPromise;

  panelPromise = (async () => {
    const created = await formula.ui.createPanel(PANEL_ID, { title: PANEL_TITLE });
    await created.webview.setHtml(html);

    if (!panelMessageDisposable) {
      panelMessageDisposable = created.webview.onDidReceiveMessage((message) => {
        if (message && message.type === "ping") {
          created.webview.postMessage({ type: "pong" }).catch((err) => {
            // eslint-disable-next-line no-console
            console.error(err);
          });
        }
      });
      context.subscriptions.push(panelMessageDisposable);
    }

    panel = created;
    context.subscriptions.push(created);
    return created;
  })();

  try {
    return await panelPromise;
  } finally {
    panelPromise = null;
  }
}

/**
 * @param {import("@formula/extension-api").ExtensionContext} context
 */
async function activate(context) {
  context.subscriptions.push(
    await formula.commands.registerCommand("sampleHello.sumSelection", async () => {
      const selection = await formula.cells.getSelection();
      const sum = sumValues(selection?.values);
      await formula.cells.setCell(2, 0, sum);
      await formula.ui.showMessage(`Sum: ${sum}`);
      return sum;
    })
  );

  context.subscriptions.push(
    await formula.commands.registerCommand("sampleHello.openPanel", async () => {
      const created = await ensurePanel(context);
      return created.id;
    })
  );

  context.subscriptions.push(
    await formula.commands.registerCommand("sampleHello.fetchText", async (url) => {
      const doFetch = typeof fetch === "function" ? fetch : formula.network.fetch;
      const response = await doFetch(String(url));
      const text = await response.text();
      await formula.ui.showMessage(`Fetched: ${text}`);
      return text;
    })
  );

  context.subscriptions.push(
    await formula.commands.registerCommand("sampleHello.copySumToClipboard", async () => {
      const sum = await getSelectionSum();
      await formula.clipboard.writeText(String(sum));
      return sum;
    })
  );

  context.subscriptions.push(
    await formula.commands.registerCommand("sampleHello.showGreeting", async () => {
      const greeting = await formula.config.get("sampleHello.greeting");
      const value = typeof greeting === "string" ? greeting : String(greeting ?? "");
      await formula.ui.showMessage(`Greeting: ${value}`);
      return value;
    })
  );

  context.subscriptions.push(
    await formula.functions.register("SAMPLEHELLO_DOUBLE", {
      handler(value) {
        return value * 2;
      }
    })
  );

  context.subscriptions.push(
    await formula.dataConnectors.register(CONNECTOR_ID, {
      async browse(_config, path) {
        if (path) return [];
        return [
          { id: "hello", name: "Hello", type: "table" },
          { id: "world", name: "World", type: "table" }
        ];
      },
      async query() {
        return {
          columns: ["id", "label"],
          rows: [
            [1, "hello"],
            [2, "world"]
          ]
        };
      },
      async testConnection() {
        return { success: true };
      }
    })
  );

  context.subscriptions.push(
    formula.events.onViewActivated(({ viewId }) => {
      if (viewId !== PANEL_ID) return;
      void ensurePanel(context).catch((error) => {
        // eslint-disable-next-line no-console
        console.error(error);
      });
    })
  );
}

module.exports = {
  activate
};
