import { optimizeMacroActions } from "./optimizeActions.js";

export class MacroRecorder {
  constructor(workbook) {
    this.workbook = workbook;
    this.recording = false;
    this.actions = [];
    this.unsubscribes = [];
  }

  start() {
    if (this.recording) return;
    this.recording = true;
    this.actions = [];

    this.unsubscribes.push(
      this.workbook.events.on("cellChanged", (evt) => {
        if (!this.recording) return;
        if (evt.values.length === 1 && evt.values[0].length === 1 && !evt.address.includes(":")) {
          this.actions.push({
            type: "setCellValue",
            sheetName: evt.sheetName,
            address: evt.address,
            value: evt.values[0][0],
          });
          return;
        }

        this.actions.push({
          type: "setRangeValues",
          sheetName: evt.sheetName,
          address: evt.address,
          values: evt.values,
        });
      }),
    );

    this.unsubscribes.push(
      this.workbook.events.on("selectionChanged", (evt) => {
        if (!this.recording) return;
        this.actions.push({ type: "setSelection", sheetName: evt.sheetName, address: evt.address });
      }),
    );

    this.unsubscribes.push(
      this.workbook.events.on("formatChanged", (evt) => {
        if (!this.recording) return;
        this.actions.push({
          type: "setFormat",
          sheetName: evt.sheetName,
          address: evt.address,
          format: evt.format,
        });
      }),
    );
  }

  stop() {
    if (!this.recording) return [];
    this.recording = false;
    for (const unsub of this.unsubscribes) unsub();
    this.unsubscribes = [];
    return this.getOptimizedActions();
  }

  getRawActions() {
    return [...this.actions];
  }

  getOptimizedActions() {
    return optimizeMacroActions(this.actions);
  }
}
