import type { RibbonTabDefinition } from "../ribbonSchema.js";

export const dataTab: RibbonTabDefinition = {
  id: "data",
  label: "Data",
  groups: [
    {
      id: "data.getTransform",
      label: "Get & Transform Data",
      buttons: [
        {
          id: "data.getTransform.getData",
          label: "Get Data",
          ariaLabel: "Get Data",
          iconId: "download",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "data.getTransform.getData.fromFile", label: "From File", ariaLabel: "From File", iconId: "file" },
            { id: "data.getTransform.getData.fromDatabase", label: "From Database", ariaLabel: "From Database", iconId: "folderOpen" },
            { id: "data.getTransform.getData.fromAzure", label: "From Azure", ariaLabel: "From Azure", iconId: "cloud" },
            { id: "data.getTransform.getData.fromOnlineServices", label: "From Online Services", ariaLabel: "From Online Services", iconId: "globe" },
            { id: "data.getTransform.getData.fromOtherSources", label: "From Other Sources", ariaLabel: "From Other Sources", iconId: "plus" },
          ],
        },
        { id: "data.getTransform.recentSources", label: "Recent Sources", ariaLabel: "Recent Sources", iconId: "clock", kind: "dropdown" },
        { id: "data.getTransform.existingConnections", label: "Existing Connections", ariaLabel: "Existing Connections", iconId: "link", kind: "dropdown" },
      ],
    },
    {
      id: "data.queriesConnections",
      label: "Queries & Connections",
      buttons: [
        {
          id: "data.queriesConnections.refreshAll",
          label: "Refresh All",
          ariaLabel: "Refresh All",
          iconId: "refresh",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "data.queriesConnections.refreshAll", label: "Refresh All", ariaLabel: "Refresh All", iconId: "refresh" },
            { id: "data.queriesConnections.refreshAll.refresh", label: "Refresh", ariaLabel: "Refresh", iconId: "refresh" },
            { id: "data.queriesConnections.refreshAll.refreshAllConnections", label: "Refresh All Connections", ariaLabel: "Refresh All Connections", iconId: "link" },
            { id: "data.queriesConnections.refreshAll.refreshAllQueries", label: "Refresh All Queries", ariaLabel: "Refresh All Queries", iconId: "folderOpen" },
          ],
        },
        { id: "data.queriesConnections.queriesConnections", label: "Queries & Connections", ariaLabel: "Queries and Connections", iconId: "folderOpen", kind: "toggle", defaultPressed: false },
        { id: "data.queriesConnections.properties", label: "Properties", ariaLabel: "Properties", iconId: "settings", kind: "dropdown" },
      ],
    },
    {
      id: "data.sortFilter",
      label: "Sort & Filter",
      buttons: [
        { id: "data.sortFilter.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", iconId: "sort" },
        { id: "data.sortFilter.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", iconId: "sort" },
        {
          id: "data.sortFilter.sort",
          label: "Sort",
          ariaLabel: "Sort",
          iconId: "sort",
          kind: "dropdown",
          menuItems: [
            { id: "data.sortFilter.sort.customSort", label: "Custom Sort…", ariaLabel: "Custom Sort", iconId: "settings" },
            { id: "data.sortFilter.sort.sortAtoZ", label: "Sort A to Z", ariaLabel: "Sort A to Z", iconId: "sort" },
            { id: "data.sortFilter.sort.sortZtoA", label: "Sort Z to A", ariaLabel: "Sort Z to A", iconId: "sort" },
          ],
        },
        { id: "data.sortFilter.filter", label: "Filter", ariaLabel: "Filter", iconId: "filter", kind: "toggle" },
        { id: "data.sortFilter.clear", label: "Clear", ariaLabel: "Clear", iconId: "close" },
        { id: "data.sortFilter.reapply", label: "Reapply", ariaLabel: "Reapply", iconId: "refresh" },
        {
          id: "data.sortFilter.advanced",
          label: "Advanced",
          ariaLabel: "Advanced",
          iconId: "settings",
          kind: "dropdown",
          menuItems: [
            { id: "data.sortFilter.advanced.advancedFilter", label: "Advanced Filter…", ariaLabel: "Advanced Filter", iconId: "settings" },
            { id: "data.sortFilter.advanced.clearFilter", label: "Clear Filter", ariaLabel: "Clear Filter", iconId: "close" },
          ],
        },
      ],
    },
    {
      id: "data.dataTools",
      label: "Data Tools",
      buttons: [
        {
          id: "data.dataTools.textToColumns",
          label: "Text to Columns",
          ariaLabel: "Text to Columns",
          iconId: "insertColumns",
          kind: "dropdown",
          menuItems: [
            { id: "data.dataTools.textToColumns", label: "Text to Columns…", ariaLabel: "Text to Columns", iconId: "insertColumns" },
            { id: "data.dataTools.textToColumns.reapply", label: "Reapply", ariaLabel: "Reapply", iconId: "refresh" },
          ],
        },
        { id: "data.dataTools.flashFill", label: "Flash Fill", ariaLabel: "Flash Fill", iconId: "lightning" },
        {
          id: "data.dataTools.removeDuplicates",
          label: "Remove Duplicates",
          ariaLabel: "Remove Duplicates",
          iconId: "trash",
          kind: "dropdown",
          menuItems: [
            { id: "data.dataTools.removeDuplicates", label: "Remove Duplicates…", ariaLabel: "Remove Duplicates", iconId: "trash" },
            { id: "data.dataTools.removeDuplicates.advanced", label: "Advanced…", ariaLabel: "Advanced", iconId: "settings" },
          ],
        },
        {
          id: "data.dataTools.dataValidation",
          label: "Data Validation",
          ariaLabel: "Data Validation",
          iconId: "check",
          kind: "dropdown",
          menuItems: [
            { id: "data.dataTools.dataValidation", label: "Data Validation…", ariaLabel: "Data Validation", iconId: "check" },
            { id: "data.dataTools.dataValidation.circleInvalid", label: "Circle Invalid Data", ariaLabel: "Circle Invalid Data", iconId: "warning" },
            { id: "data.dataTools.dataValidation.clearCircles", label: "Clear Validation Circles", ariaLabel: "Clear Validation Circles", iconId: "close" },
          ],
        },
        {
          id: "data.dataTools.consolidate",
          label: "Consolidate",
          ariaLabel: "Consolidate",
          iconId: "puzzle",
          kind: "dropdown",
          menuItems: [{ id: "data.dataTools.consolidate", label: "Consolidate…", ariaLabel: "Consolidate", iconId: "puzzle" }],
        },
        {
          id: "data.dataTools.relationships",
          label: "Relationships",
          ariaLabel: "Relationships",
          iconId: "link",
          kind: "dropdown",
          menuItems: [
            { id: "data.dataTools.relationships", label: "Relationships…", ariaLabel: "Relationships", iconId: "link" },
            { id: "data.dataTools.relationships.manage", label: "Manage Relationships…", ariaLabel: "Manage Relationships", iconId: "settings" },
          ],
        },
        {
          id: "data.dataTools.manageDataModel",
          label: "Manage Data Model",
          ariaLabel: "Manage Data Model",
          iconId: "puzzle",
          kind: "dropdown",
          menuItems: [
            { id: "data.dataTools.manageDataModel", label: "Manage Data Model", ariaLabel: "Manage Data Model", iconId: "puzzle" },
            { id: "data.dataTools.manageDataModel.addToDataModel", label: "Add to Data Model", ariaLabel: "Add to Data Model", iconId: "plus" },
          ],
        },
      ],
    },
    {
      id: "data.forecast",
      label: "Forecast",
      buttons: [
        {
          id: "data.forecast.whatIfAnalysis",
          label: "What-If Analysis",
          ariaLabel: "What-If Analysis",
          iconId: "help",
          kind: "dropdown",
          size: "large",
          menuItems: [
            { id: "data.forecast.whatIfAnalysis.scenarioManager", label: "Scenario Manager…", ariaLabel: "Scenario Manager", iconId: "settings" },
            { id: "data.forecast.whatIfAnalysis.goalSeek", label: "Goal Seek…", ariaLabel: "Goal Seek", iconId: "find" },
            { id: "data.forecast.whatIfAnalysis.monteCarlo", label: "Monte Carlo…", ariaLabel: "Monte Carlo", iconId: "chart" },
            { id: "data.forecast.whatIfAnalysis.dataTable", label: "Data Table…", ariaLabel: "Data Table", iconId: "gridlines" },
          ],
        },
        {
          id: "data.forecast.forecastSheet",
          label: "Forecast Sheet",
          ariaLabel: "Forecast Sheet",
          iconId: "chart",
          kind: "dropdown",
          menuItems: [
            { id: "data.forecast.forecastSheet", label: "Forecast Sheet…", ariaLabel: "Forecast Sheet", iconId: "chart" },
            { id: "data.forecast.forecastSheet.options", label: "Options…", ariaLabel: "Forecast Options", iconId: "settings" },
          ],
        },
      ],
    },
    {
      id: "data.outline",
      label: "Outline",
      buttons: [
        {
          id: "data.outline.group",
          label: "Group",
          ariaLabel: "Group",
          iconId: "plus",
          kind: "dropdown",
          menuItems: [
            { id: "data.outline.group.group", label: "Group…", ariaLabel: "Group", iconId: "plus" },
            { id: "data.outline.group.groupSelection", label: "Group Selection", ariaLabel: "Group Selection", iconId: "gridlines" },
          ],
        },
        {
          id: "data.outline.ungroup",
          label: "Ungroup",
          ariaLabel: "Ungroup",
          iconId: "minus",
          kind: "dropdown",
          menuItems: [
            { id: "data.outline.ungroup.ungroup", label: "Ungroup…", ariaLabel: "Ungroup", iconId: "minus" },
            { id: "data.outline.ungroup.clearOutline", label: "Clear Outline", ariaLabel: "Clear Outline", iconId: "close" },
          ],
        },
        {
          id: "data.outline.subtotal",
          label: "Subtotal",
          ariaLabel: "Subtotal",
          iconId: "autoSum",
          kind: "dropdown",
          menuItems: [{ id: "data.outline.subtotal", label: "Subtotal…", ariaLabel: "Subtotal", iconId: "autoSum" }],
        },
        { id: "data.outline.showDetail", label: "Show Detail", ariaLabel: "Show Detail", iconId: "plus" },
        { id: "data.outline.hideDetail", label: "Hide Detail", ariaLabel: "Hide Detail", iconId: "minus" },
      ],
    },
    {
      id: "data.dataTypes",
      label: "Data Types",
      buttons: [
        { id: "data.dataTypes.stocks", label: "Stocks", ariaLabel: "Stocks", iconId: "chart", kind: "dropdown", size: "large" },
        { id: "data.dataTypes.geography", label: "Geography", ariaLabel: "Geography", iconId: "globe", kind: "dropdown", size: "large" },
      ],
    },
  ],
};
