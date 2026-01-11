/**
 * Minimal, dependency-free "in-app" marketplace panel.
 *
 * The real desktop app will render this inside the application's panel system
 * (Tauri/WebView). This module focuses on wiring: search → install/update/uninstall.
 */

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "className") node.className = value;
    else if (key.startsWith("on") && typeof value === "function") node.addEventListener(key.slice(2).toLowerCase(), value);
    else node.setAttribute(key, String(value));
  }
  for (const child of children) node.append(child);
  return node;
}

async function renderSearchResults({ container, marketplaceClient, extensionManager, extensionHostManager, query }) {
  container.textContent = "Searching…";
  const results = await marketplaceClient.search({ q: query, limit: 25, offset: 0 });

  const list = el("div", { className: "marketplace-results" });

  for (const item of results.results) {
    const installed = await extensionManager.getInstalled(item.id);
    const row = el("div", { className: "marketplace-result" }, [
      el("div", { className: "title" }, [document.createTextNode(`${item.displayName} (${item.id})`)]),
      el("div", { className: "desc" }, [document.createTextNode(item.description || "")]),
    ]);

    const actions = el("div", { className: "actions" });
    if (!installed) {
      actions.append(
        el("button", {
          onClick: async () => {
            actions.textContent = "Installing…";
            await extensionManager.install(item.id);
            if (extensionHostManager) {
              await extensionHostManager.reloadExtension(item.id);
            }
            actions.textContent = "Installed";
          },
        }, [document.createTextNode("Install")]),
      );
    } else {
      actions.append(
        el("button", {
          onClick: async () => {
            actions.textContent = "Uninstalling…";
            if (extensionHostManager) {
              await extensionHostManager.unloadExtension(item.id);
            }
            await extensionManager.uninstall(item.id);
            actions.textContent = "Uninstalled";
          },
        }, [document.createTextNode("Uninstall")]),
      );

      actions.append(
        el("button", {
          onClick: async () => {
            actions.textContent = "Checking…";
            const updates = await extensionManager.checkForUpdates();
            const update = updates.find((u) => u.id === item.id);
            if (!update) {
              actions.textContent = "Up to date";
              return;
            }
            actions.textContent = `Updating to ${update.latestVersion}…`;
            // Terminate the running extension before mutating its install directory.
            // This avoids worker threads reading partially-updated files.
            if (extensionHostManager) {
              await extensionHostManager.unloadExtension(item.id);
            }
            await extensionManager.update(item.id);
            if (extensionHostManager) {
              await extensionHostManager.reloadExtension(item.id);
            }
            actions.textContent = "Updated";
          },
        }, [document.createTextNode("Update")]),
      );
    }

    row.append(actions);
    list.append(row);
  }

  container.textContent = "";
  container.append(list);
}

export function createMarketplacePanel({ container, marketplaceClient, extensionManager, extensionHostManager }) {
  const queryInput = el("input", { type: "search", placeholder: "Search extensions…" });
  const resultsContainer = el("div", { className: "results" });

  const searchButton = el(
    "button",
    {
      onClick: async () => {
        await renderSearchResults({
          container: resultsContainer,
          marketplaceClient,
          extensionManager,
          extensionHostManager,
          query: queryInput.value,
        });
      },
    },
    [document.createTextNode("Search")],
  );

  container.append(el("div", { className: "marketplace-panel" }, [queryInput, searchButton, resultsContainer]));

  return {
    dispose() {
      container.textContent = "";
    },
  };
}
