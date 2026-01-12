const test = require("node:test");
const assert = require("node:assert/strict");

const { ExtensionHost } = require("../src");

test("ExtensionHost.dispose unsubscribes from spreadsheet event hooks", async () => {
  let selectionDisposed = false;
  let cellDisposed = false;
  let sheetDisposed = false;

  const host = new ExtensionHost({
    spreadsheet: {
      onSelectionChanged() {
        return () => {
          selectionDisposed = true;
        };
      },
      onCellChanged() {
        return { dispose: () => {
          cellDisposed = true;
        } };
      },
      onSheetActivated() {
        return () => {
          sheetDisposed = true;
        };
      },
    },
  });

  await host.dispose();

  assert.equal(selectionDisposed, true);
  assert.equal(cellDisposed, true);
  assert.equal(sheetDisposed, true);
});

