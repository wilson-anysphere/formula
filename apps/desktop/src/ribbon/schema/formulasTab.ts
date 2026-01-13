import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const formulasTab: RibbonTabDefinition = {
  id: "formulas",
  label: "Formulas",
  groups: [
    {
      id: "formulas.functionLibrary",
      label: "Function Library",
      buttons: [
         { id: "formulas.functionLibrary.insertFunction", label: "Insert Function", ariaLabel: "Insert Function", iconId: "function", kind: "dropdown", size: "large" },
         {
           id: "formulas.functionLibrary.autoSum",
           label: "AutoSum",
           ariaLabel: "AutoSum",
           iconId: "autoSum",
           kind: "dropdown",
           menuItems: [
             { id: "formulas.functionLibrary.autoSum.sum", label: "Sum", ariaLabel: "Sum", iconId: "autoSum" },
             { id: "formulas.functionLibrary.autoSum.average", label: "Average", ariaLabel: "Average", iconId: "divide" },
             { id: "formulas.functionLibrary.autoSum.countNumbers", label: "Count Numbers", ariaLabel: "Count Numbers", iconId: "hash" },
             { id: "formulas.functionLibrary.autoSum.max", label: "Max", ariaLabel: "Max", iconId: "arrowUp" },
             { id: "formulas.functionLibrary.autoSum.min", label: "Min", ariaLabel: "Min", iconId: "arrowDown" },
             { id: "formulas.functionLibrary.autoSum.moreFunctions", label: "More Functions…", ariaLabel: "More Functions", iconId: "function" },
           ],
        },
        { id: "formulas.functionLibrary.recentlyUsed", label: "Recently Used", ariaLabel: "Recently Used", iconId: "clock", kind: "dropdown" },
        { id: "formulas.functionLibrary.financial", label: "Financial", ariaLabel: "Financial", iconId: "currency", kind: "dropdown" },
        { id: "formulas.functionLibrary.logical", label: "Logical", ariaLabel: "Logical", iconId: "function", kind: "dropdown" },
        { id: "formulas.functionLibrary.text", label: "Text", ariaLabel: "Text", iconId: "fontSize", kind: "dropdown" },
        { id: "formulas.functionLibrary.dateTime", label: "Date & Time", ariaLabel: "Date and Time", iconId: "calendar", kind: "dropdown" },
        { id: "formulas.functionLibrary.lookupReference", label: "Lookup & Reference", ariaLabel: "Lookup and Reference", iconId: "search", kind: "dropdown" },
        { id: "formulas.functionLibrary.mathTrig", label: "Math & Trig", ariaLabel: "Math and Trig", iconId: "pi", kind: "dropdown" },
        { id: "formulas.functionLibrary.moreFunctions", label: "More Functions", ariaLabel: "More Functions", iconId: "plus", kind: "dropdown" },
      ],
    },
    {
      id: "formulas.definedNames",
      label: "Defined Names",
      buttons: [
        { id: "formulas.definedNames.nameManager", label: "Name Manager", ariaLabel: "Name Manager", iconId: "tag", kind: "dropdown", size: "large" },
        { id: "formulas.definedNames.defineName", label: "Define Name", ariaLabel: "Define Name", iconId: "plus", kind: "dropdown" },
        { id: "formulas.definedNames.useInFormula", label: "Use in Formula", ariaLabel: "Use in Formula", iconId: "function", kind: "dropdown" },
        { id: "formulas.definedNames.createFromSelection", label: "Create from Selection", ariaLabel: "Create from Selection", iconId: "gridlines", kind: "dropdown" },
      ],
    },
    {
      id: "formulas.formulaAuditing",
      label: "Formula Auditing",
      buttons: [
        { id: "formulas.formulaAuditing.tracePrecedents", label: "Trace Precedents", ariaLabel: "Trace Precedents", iconId: "arrowLeft", size: "small" },
        { id: "formulas.formulaAuditing.traceDependents", label: "Trace Dependents", ariaLabel: "Trace Dependents", iconId: "arrowRight", size: "small" },
        { id: "formulas.formulaAuditing.removeArrows", label: "Remove Arrows", ariaLabel: "Remove Arrows", iconId: "close", kind: "dropdown", size: "small" },
        { id: "view.toggleShowFormulas", label: "Show Formulas", ariaLabel: "Show Formulas", iconId: "function", kind: "toggle", size: "small" },
        { id: "formulas.formulaAuditing.errorChecking", label: "Error Checking", ariaLabel: "Error Checking", iconId: "warning", kind: "dropdown", size: "small" },
        { id: "formulas.formulaAuditing.evaluateFormula", label: "Evaluate Formula", ariaLabel: "Evaluate Formula", iconId: "autoSum", kind: "dropdown", size: "small" },
        { id: "formulas.formulaAuditing.watchWindow", label: "Watch Window", ariaLabel: "Watch Window", iconId: "eye", kind: "dropdown", size: "small" },
      ],
    },
    {
      id: "formulas.calculation",
      label: "Calculation",
      buttons: [
        { id: "formulas.calculation.calculationOptions", label: "Calculation Options", ariaLabel: "Calculation Options", iconId: "settings", kind: "dropdown", size: "large" },
        { id: "formulas.calculation.calculateNow", label: "Calculate Now", ariaLabel: "Calculate Now", iconId: "refresh", size: "small" },
        { id: "formulas.calculation.calculateSheet", label: "Calculate Sheet", ariaLabel: "Calculate Sheet", iconId: "refresh", size: "small" },
      ],
    },
    {
      id: "formulas.solutions",
      label: "Solutions",
      buttons: [
        { id: "formulas.solutions.solver", label: "Solver", ariaLabel: "Solver", iconId: "puzzle", kind: "dropdown", size: "large" },
        { id: "formulas.solutions.analysisToolPak", label: "Analysis ToolPak", ariaLabel: "Analysis ToolPak", iconId: "settings", kind: "dropdown" },
      ],
    },
  ],
};
