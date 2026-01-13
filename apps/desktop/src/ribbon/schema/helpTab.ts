import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const helpTab: RibbonTabDefinition = {
  id: "help",
  label: "Help",
  groups: [
    {
      id: "help.support",
      label: "Support",
      buttons: [
        { id: "help.support.help", label: "Help", ariaLabel: "Help", iconId: "help", kind: "dropdown", size: "large" },
        { id: "help.support.training", label: "Training", ariaLabel: "Training", iconId: "help", kind: "dropdown" },
        { id: "help.support.contactSupport", label: "Contact Support", ariaLabel: "Contact Support", iconId: "help", kind: "dropdown" },
        { id: "help.support.feedback", label: "Feedback", ariaLabel: "Feedback", iconId: "edit", kind: "dropdown" },
      ],
    },
  ],
};
