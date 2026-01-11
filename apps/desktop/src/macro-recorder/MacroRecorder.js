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

    const isFormulaInput = (value) => {
      if (typeof value !== "string") return false;
      if (value.startsWith("'")) return false;
      const trimmed = value.trimStart();
      return trimmed.startsWith("=") && trimmed.length > 1;
    };

    this.unsubscribes.push(
      this.workbook.events.on("cellChanged", (evt) => {
        if (!this.recording) return;
        if (evt.values.length === 1 && evt.values[0].length === 1 && !evt.address.includes(":")) {
          const value = evt.values[0][0];
          if (isFormulaInput(value)) {
            this.actions.push({
              type: "setCellFormula",
              sheetName: evt.sheetName,
              address: evt.address,
              formula: value,
            });
          } else {
            this.actions.push({
              type: "setCellValue",
              sheetName: evt.sheetName,
              address: evt.address,
              value,
            });
          }
          return;
        }

        const allFormulas = evt.values.every((row) => row.every((value) => isFormulaInput(value)));
        if (allFormulas) {
          this.actions.push({
            type: "setRangeFormulas",
            sheetName: evt.sheetName,
            address: evt.address,
            formulas: evt.values,
          });
        } else {
          this.actions.push({
            type: "setRangeValues",
            sheetName: evt.sheetName,
            address: evt.address,
            values: evt.values,
          });
        }
      }),
    );

    this.unsubscribes.push(
      this.workbook.events.on("formulaChanged", (evt) => {
        if (!this.recording) return;
        if (evt.formulas.length === 1 && evt.formulas[0].length === 1 && !evt.address.includes(":")) {
          this.actions.push({
            type: "setCellFormula",
            sheetName: evt.sheetName,
            address: evt.address,
            formula: evt.formulas[0][0],
          });
          return;
        }

        this.actions.push({
          type: "setRangeFormulas",
          sheetName: evt.sheetName,
          address: evt.address,
          formulas: evt.formulas,
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
