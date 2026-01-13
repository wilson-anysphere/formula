import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const developerTab: RibbonTabDefinition = {
  id: "developer",
  label: "Developer",
  groups: [
    {
      id: "developer.code",
      label: "Code",
      buttons: [
        {
          id: "developer.code.visualBasic",
          label: "Visual Basic",
          ariaLabel: "Visual Basic",
          iconId: "code",
          size: "large",
          testId: "ribbon-developer-visual-basic",
        },
        {
          id: "developer.code.macros",
          label: "Macros",
          ariaLabel: "Macros",
          iconId: "file",
          kind: "dropdown",
          size: "large",
          testId: "ribbon-developer-macros",
          menuItems: [
            { id: "developer.code.macros", label: "Macros…", ariaLabel: "Macros", iconId: "file", testId: "ribbon-developer-macros-open" },
            { id: "developer.code.macros.run", label: "Run…", ariaLabel: "Run Macro", iconId: "play", testId: "ribbon-developer-macros-run" },
            { id: "developer.code.macros.edit", label: "Edit…", ariaLabel: "Edit Macro", iconId: "edit", testId: "ribbon-developer-macros-edit" },
          ],
        },
        {
          id: "developer.code.recordMacro",
          label: "Record Macro",
          ariaLabel: "Record Macro",
          iconId: "record",
          kind: "dropdown",
          testId: "ribbon-developer-record-macro",
          menuItems: [
            { id: "developer.code.recordMacro", label: "Record Macro…", ariaLabel: "Record Macro", iconId: "record", testId: "ribbon-developer-record-macro-start" },
            { id: "developer.code.recordMacro.stop", label: "Stop Recording", ariaLabel: "Stop Recording", iconId: "stop", testId: "ribbon-developer-record-macro-stop" },
          ],
        },
        { id: "developer.code.useRelativeReferences", label: "Use Relative References", ariaLabel: "Use Relative References", iconId: "pin", kind: "toggle" },
        {
          id: "developer.code.macroSecurity",
          label: "Macro Security",
          ariaLabel: "Macro Security",
          iconId: "lock",
          kind: "dropdown",
          testId: "ribbon-developer-macro-security",
          menuItems: [
            { id: "developer.code.macroSecurity", label: "Macro Security…", ariaLabel: "Macro Security", iconId: "lock", testId: "ribbon-developer-macro-security-open" },
            { id: "developer.code.macroSecurity.trustCenter", label: "Trust Center…", ariaLabel: "Trust Center", iconId: "lock", testId: "ribbon-developer-macro-security-trust-center" },
          ],
        },
      ],
    },
    {
      id: "developer.addins",
      label: "Add-ins",
      buttons: [
        {
          id: "developer.addins.addins",
          label: "Add-ins",
          ariaLabel: "Add-ins",
          iconId: "puzzle",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "developer.addins.addins.excelAddins", label: "Excel Add-ins…", ariaLabel: "Excel Add-ins", iconId: "puzzle" },
            { id: "developer.addins.addins.browse", label: "Browse…", ariaLabel: "Browse Add-ins", iconId: "folderOpen" },
            { id: "developer.addins.addins.manage", label: "Manage…", ariaLabel: "Manage Add-ins", iconId: "settings" },
          ],
        },
        {
          id: "developer.addins.comAddins",
          label: "COM Add-ins",
          ariaLabel: "COM Add-ins",
          iconId: "puzzle",
          kind: "dropdown",
          menuItems: [{ id: "developer.addins.comAddins", label: "COM Add-ins…", ariaLabel: "COM Add-ins", iconId: "puzzle" }],
        },
      ],
    },
    {
      id: "developer.controls",
      label: "Controls",
      buttons: [
        {
          id: "developer.controls.insert",
          label: "Insert",
          ariaLabel: "Insert Control",
          iconId: "plus",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "developer.controls.insert.button", label: "Button", ariaLabel: "Insert Button", iconId: "check" },
            { id: "developer.controls.insert.checkbox", label: "Check Box", ariaLabel: "Insert Check Box", iconId: "check" },
            { id: "developer.controls.insert.combobox", label: "Combo Box", ariaLabel: "Insert Combo Box", iconId: "arrowDown" },
            { id: "developer.controls.insert.listbox", label: "List Box", ariaLabel: "Insert List Box", iconId: "menu" },
            { id: "developer.controls.insert.scrollbar", label: "Scroll Bar", ariaLabel: "Insert Scroll Bar", iconId: "arrowUpDown" },
            { id: "developer.controls.insert.spinButton", label: "Spin Button", ariaLabel: "Insert Spin Button", iconId: "arrowUpDown" },
          ],
        },
        { id: "developer.controls.designMode", label: "Design Mode", ariaLabel: "Design Mode", iconId: "settings", kind: "toggle" },
        {
          id: "developer.controls.properties",
          label: "Properties",
          ariaLabel: "Properties",
          iconId: "settings",
          kind: "dropdown",
          menuItems: [
            { id: "developer.controls.properties", label: "Properties…", ariaLabel: "Properties", iconId: "settings" },
            { id: "developer.controls.properties.viewProperties", label: "View Properties", ariaLabel: "View Properties", iconId: "eye" },
          ],
        },
        { id: "developer.controls.viewCode", label: "View Code", ariaLabel: "View Code", iconId: "code" },
        { id: "developer.controls.runDialog", label: "Run Dialog", ariaLabel: "Run Dialog", iconId: "play" },
      ],
    },
    {
      id: "developer.xml",
      label: "XML",
      buttons: [
        {
          id: "developer.xml.source",
          label: "Source",
          ariaLabel: "XML Source",
          iconId: "code",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "developer.xml.source", label: "XML Source", ariaLabel: "XML Source", iconId: "code" },
            { id: "developer.xml.source.refresh", label: "Refresh XML Data", ariaLabel: "Refresh XML Data", iconId: "refresh" },
          ],
        },
        {
          id: "developer.xml.mapProperties",
          label: "Map Properties",
          ariaLabel: "Map Properties",
          iconId: "globe",
          kind: "dropdown",
          menuItems: [{ id: "developer.xml.mapProperties", label: "Map Properties…", ariaLabel: "Map Properties", iconId: "globe" }],
        },
        { id: "developer.xml.import", label: "Import", ariaLabel: "Import XML", iconId: "download" },
        { id: "developer.xml.export", label: "Export", ariaLabel: "Export XML", iconId: "upload" },
        { id: "developer.xml.refreshData", label: "Refresh Data", ariaLabel: "Refresh Data", iconId: "refresh" },
      ],
    },
  ],
};
