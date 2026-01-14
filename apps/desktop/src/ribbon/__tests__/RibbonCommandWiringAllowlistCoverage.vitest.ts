import { describe, expect, it } from "vitest";

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { registerDesktopCommands } from "../../commands/registerDesktopCommands.js";
import { registerDataQueriesCommands } from "../../commands/registerDataQueriesCommands.js";
import { registerFormatPainterCommand } from "../../commands/formatPainterCommand.js";
import { registerRibbonMacroCommands } from "../../commands/registerRibbonMacroCommands.js";
import { createDefaultLayout } from "../../layout/layoutState.js";

import type { RibbonSchema } from "../ribbonSchema";
import { defaultRibbonSchema } from "../ribbonSchema";

function collectRibbonCommandIds(schema: RibbonSchema): string[] {
  const ids = new Set<string>();
  for (const tab of schema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        ids.add(button.id);
        for (const item of button.menuItems ?? []) {
          ids.add(item.id);
        }
      }
    }
  }
  return Array.from(ids).sort();
}

/**
 * All ribbon command ids that are expected to be explicitly handled by the desktop app.
 *
 * This list intentionally includes both:
 * - command ids that invoke real functionality, and
 * - command ids that are explicitly handled to avoid falling back to the default
 *   `showToast(`Ribbon: ${commandId}`)` implementation.
 *
 * If you add a new command id to `defaultRibbonSchema`, it must be added here (when
 * wired) or to `knownUnimplementedCommandIds` (when intentionally left unwired).
 */
const implementedCommandIds: string[] = [
  "ai.inlineEdit",
  "audit.toggleDependents",
  "audit.togglePrecedents",
  "audit.toggleTransitive",
  "clipboard.copy",
  "clipboard.cut",
  "clipboard.paste",
  "clipboard.pasteSpecial",
  "clipboard.pasteSpecial.formats",
  "clipboard.pasteSpecial.formulas",
  "clipboard.pasteSpecial.transpose",
  "clipboard.pasteSpecial.values",
  "comments.addComment",
  "comments.togglePanel",
  "data.queriesConnections.queriesConnections",
  "data.queriesConnections.refreshAll",
  "data.queriesConnections.refreshAll.refresh",
  "data.queriesConnections.refreshAll.refreshAllConnections",
  "data.queriesConnections.refreshAll.refreshAllQueries",
  "data.forecast.whatIfAnalysis.goalSeek",
  "data.forecast.whatIfAnalysis.scenarioManager",
  "data.sortFilter.advanced.clearFilter",
  "data.sortFilter.clear",
  "data.sortFilter.filter",
  "data.sortFilter.reapply",
  "data.sortFilter.sort.customSort",
  "data.sortFilter.sort.sortAtoZ",
  "data.sortFilter.sort.sortZtoA",
  "data.sortFilter.sortAtoZ",
  "data.sortFilter.sortZtoA",
  "developer.code.macroSecurity",
  "developer.code.macroSecurity.trustCenter",
  "developer.code.macros",
  "developer.code.macros.edit",
  "developer.code.macros.run",
  "developer.code.recordMacro",
  "developer.code.recordMacro.stop",
  "developer.code.useRelativeReferences",
  "developer.code.visualBasic",
  "edit.autoSum",
  "edit.clearContents",
  "edit.fillDown",
  "edit.fillRight",
  "edit.find",
  "edit.replace",
  "file.export.changeFileType.csv",
  "file.export.changeFileType.pdf",
  "file.export.changeFileType.tsv",
  "file.export.changeFileType.xlsx",
  "file.export.createPdf",
  "file.export.export.csv",
  "file.export.export.pdf",
  "file.export.export.xlsx",
  "file.info.manageWorkbook.branches",
  "file.info.manageWorkbook.versions",
  "file.new.blankWorkbook",
  "file.new.new",
  "file.open.open",
  "file.options.close",
  "file.print.pageSetup",
  "file.print.pageSetup.margins",
  "file.print.pageSetup.printTitles",
  "file.print.print",
  "file.print.printPreview",
  "file.save.autoSave",
  "file.save.save",
  "file.save.saveAs",
  "file.save.saveAs.copy",
  "file.save.saveAs.download",
  "format.alignBottom",
  "format.alignCenter",
  "format.alignLeft",
  "format.alignMiddle",
  "format.alignRight",
  "format.alignTop",
  "format.borders.all",
  "format.borders.bottom",
  "format.borders.left",
  "format.borders.none",
  "format.borders.outside",
  "format.borders.right",
  "format.borders.thickBox",
  "format.borders.top",
  "format.clearAll",
  "format.clearContents",
  "format.clearFormats",
  "format.decreaseFontSize",
  "format.decreaseIndent",
  "format.fillColor.blue",
  "format.fillColor.green",
  "format.fillColor.lightGray",
  "format.fillColor.moreColors",
  "format.fillColor.none",
  "format.fillColor.red",
  "format.fillColor.yellow",
  "format.fontColor.automatic",
  "format.fontColor.black",
  "format.fontColor.blue",
  "format.fontColor.green",
  "format.fontColor.moreColors",
  "format.fontColor.red",
  "format.fontName.arial",
  "format.fontName.calibri",
  "format.fontName.courier",
  "format.fontName.times",
  "format.fontSize.10",
  "format.fontSize.11",
  "format.fontSize.12",
  "format.fontSize.14",
  "format.fontSize.16",
  "format.fontSize.18",
  "format.fontSize.20",
  "format.fontSize.24",
  "format.fontSize.28",
  "format.fontSize.36",
  "format.fontSize.48",
  "format.fontSize.72",
  "format.fontSize.8",
  "format.fontSize.9",
  "format.increaseFontSize",
  "format.increaseIndent",
  "format.numberFormat.accounting",
  "format.numberFormat.accounting.eur",
  "format.numberFormat.accounting.gbp",
  "format.numberFormat.accounting.jpy",
  "format.numberFormat.accounting.usd",
  "format.numberFormat.commaStyle",
  "format.numberFormat.currency",
  "format.numberFormat.decreaseDecimal",
  "format.numberFormat.fraction",
  "format.numberFormat.general",
  "format.numberFormat.increaseDecimal",
  "format.numberFormat.longDate",
  "format.numberFormat.number",
  "format.numberFormat.percent",
  "format.numberFormat.scientific",
  "format.numberFormat.shortDate",
  "format.numberFormat.text",
  "format.numberFormat.time",
  "format.openAlignmentDialog",
  "format.openFormatCells",
  "format.textRotation.angleClockwise",
  "format.textRotation.angleCounterclockwise",
  "format.textRotation.rotateDown",
  "format.textRotation.rotateUp",
  "format.textRotation.verticalText",
  "format.toggleBold",
  "format.toggleFormatPainter",
  "format.toggleItalic",
  "format.toggleStrikethrough",
  "format.toggleUnderline",
  "format.toggleWrapText",
  "formulas.formulaAuditing.removeArrows",
  "formulas.formulaAuditing.traceDependents",
  "formulas.formulaAuditing.tracePrecedents",
  "formulas.solutions.solver",
  "home.alignment.mergeCenter.mergeAcross",
  "home.alignment.mergeCenter.mergeCells",
  "home.alignment.mergeCenter.mergeCenter",
  "home.alignment.mergeCenter.unmergeCells",
  "home.cells.delete.deleteCells",
  "home.cells.delete.deleteSheet",
  "home.cells.delete.deleteSheetColumns",
  "home.cells.delete.deleteSheetRows",
  "home.cells.format",
  "home.cells.format.columnWidth",
  "home.cells.format.organizeSheets",
  "home.cells.format.rowHeight",
  "home.cells.insert.insertCells",
  "home.cells.insert.insertSheet",
  "home.cells.insert.insertSheetColumns",
  "home.cells.insert.insertSheetRows",
  "home.editing.autoSum.average",
  "home.editing.autoSum.countNumbers",
  "home.editing.autoSum.max",
  "home.editing.autoSum.min",
  "home.editing.fill.left",
  "home.editing.fill.series",
  "home.editing.fill.up",
  "home.editing.sortFilter.clear",
  "home.editing.sortFilter.customSort",
  "home.editing.sortFilter.filter",
  "home.editing.sortFilter.reapply",
  "home.editing.sortFilter.sortAtoZ",
  "home.editing.sortFilter.sortZtoA",
  "home.font.subscript",
  "home.font.superscript",
  "home.number.moreFormats.custom",
  "home.styles.cellStyles.dataModel",
  "home.styles.cellStyles.goodBadNeutral",
  "home.styles.cellStyles.newStyle",
  "home.styles.cellStyles.numberFormat",
  "home.styles.cellStyles.titlesHeadings",
  "home.styles.formatAsTable",
  "home.styles.formatAsTable.dark",
  "home.styles.formatAsTable.light",
  "home.styles.formatAsTable.medium",
  "home.styles.formatAsTable.newStyle",
  "insert.illustrations.onlinePictures",
  "insert.illustrations.pictures",
  "insert.illustrations.pictures.onlinePictures",
  "insert.illustrations.pictures.stockImages",
  "insert.illustrations.pictures.thisDevice",
  "insert.tables.pivotTable.fromTableRange",
  "navigation.goTo",
  "pageLayout.arrange.bringForward",
  "pageLayout.arrange.sendBackward",
  "pageLayout.export.exportPdf",
  "pageLayout.pageSetup.margins.custom",
  "pageLayout.pageSetup.margins.narrow",
  "pageLayout.pageSetup.margins.normal",
  "pageLayout.pageSetup.margins.wide",
  "pageLayout.pageSetup.orientation.landscape",
  "pageLayout.pageSetup.orientation.portrait",
  "pageLayout.pageSetup.pageSetupDialog",
  "pageLayout.pageSetup.printArea.addTo",
  "pageLayout.pageSetup.printArea.clear",
  "pageLayout.pageSetup.printArea.set",
  "pageLayout.pageSetup.size.a4",
  "pageLayout.pageSetup.size.letter",
  "pageLayout.pageSetup.size.more",
  "pageLayout.printArea.clearPrintArea",
  "pageLayout.printArea.setPrintArea",
  "view.appearance.theme",
  "view.appearance.theme.dark",
  "view.appearance.theme.highContrast",
  "view.appearance.theme.light",
  "view.appearance.theme.system",
  "view.freezeFirstColumn",
  "view.freezePanes",
  "view.freezeTopRow",
  "view.insertPivotTable",
  "view.macros.recordMacro",
  "view.macros.recordMacro.stop",
  "view.macros.useRelativeReferences",
  "view.macros.viewMacros",
  "view.macros.viewMacros.delete",
  "view.macros.viewMacros.edit",
  "view.macros.viewMacros.run",
  "view.splitHorizontal",
  "view.splitNone",
  "view.splitVertical",
  "view.togglePanel.aiAudit",
  "view.togglePanel.aiChat",
  "view.togglePanel.branchManager",
  "view.togglePanel.dataQueries",
  "view.togglePanel.extensions",
  "view.togglePanel.macros",
  "view.togglePanel.marketplace",
  "view.togglePanel.python",
  "view.togglePanel.scriptEditor",
  "view.togglePanel.vbaMigrate",
  "view.togglePanel.versionHistory",
  "view.togglePerformanceStats",
  "view.toggleShowFormulas",
  "view.toggleSplitView",
  "view.unfreezePanes",
  "view.zoom.openPicker",
  "view.zoom.zoom",
  "view.zoom.zoom100",
  "view.zoom.zoom150",
  "view.zoom.zoom200",
  "view.zoom.zoom25",
  "view.zoom.zoom400",
  "view.zoom.zoom50",
  "view.zoom.zoom75",
  "view.zoom.zoomToSelection",
];

/**
 * All other ribbon command ids that exist in `defaultRibbonSchema` but are
 * intentionally not wired yet.
 */
const knownUnimplementedCommandIds: string[] = [
  "data.dataTools.consolidate",
  "data.dataTools.dataValidation",
  "data.dataTools.dataValidation.circleInvalid",
  "data.dataTools.dataValidation.clearCircles",
  "data.dataTools.flashFill",
  "data.dataTools.manageDataModel",
  "data.dataTools.manageDataModel.addToDataModel",
  "data.dataTools.relationships",
  "data.dataTools.relationships.manage",
  "data.dataTools.removeDuplicates",
  "data.dataTools.removeDuplicates.advanced",
  "data.dataTools.textToColumns",
  "data.dataTools.textToColumns.reapply",
  "data.dataTypes.geography",
  "data.dataTypes.stocks",
  "data.forecast.forecastSheet",
  "data.forecast.forecastSheet.options",
  "data.forecast.whatIfAnalysis",
  "data.forecast.whatIfAnalysis.dataTable",
  "data.getTransform.existingConnections",
  "data.getTransform.getData",
  "data.getTransform.getData.fromAzure",
  "data.getTransform.getData.fromDatabase",
  "data.getTransform.getData.fromFile",
  "data.getTransform.getData.fromOnlineServices",
  "data.getTransform.getData.fromOtherSources",
  "data.getTransform.recentSources",
  "data.outline.group",
  "data.outline.group.group",
  "data.outline.group.groupSelection",
  "data.outline.hideDetail",
  "data.outline.showDetail",
  "data.outline.subtotal",
  "data.outline.ungroup",
  "data.outline.ungroup.clearOutline",
  "data.outline.ungroup.ungroup",
  "data.queriesConnections.properties",
  "data.sortFilter.advanced",
  "data.sortFilter.advanced.advancedFilter",
  "data.sortFilter.sort",
  "developer.addins.addins",
  "developer.addins.addins.browse",
  "developer.addins.addins.excelAddins",
  "developer.addins.addins.manage",
  "developer.addins.comAddins",
  "developer.controls.designMode",
  "developer.controls.insert",
  "developer.controls.insert.button",
  "developer.controls.insert.checkbox",
  "developer.controls.insert.combobox",
  "developer.controls.insert.listbox",
  "developer.controls.insert.scrollbar",
  "developer.controls.insert.spinButton",
  "developer.controls.properties",
  "developer.controls.properties.viewProperties",
  "developer.controls.runDialog",
  "developer.controls.viewCode",
  "developer.xml.export",
  "developer.xml.import",
  "developer.xml.mapProperties",
  "developer.xml.refreshData",
  "developer.xml.source",
  "developer.xml.source.refresh",
  "file.export.changeFileType",
  "file.export.export",
  "file.info.inspectWorkbook",
  "file.info.inspectWorkbook.checkAccessibility",
  "file.info.inspectWorkbook.checkCompatibility",
  "file.info.inspectWorkbook.documentInspector",
  "file.info.manageWorkbook",
  "file.info.manageWorkbook.properties",
  "file.info.manageWorkbook.recoverUnsaved",
  "file.info.protectWorkbook",
  "file.info.protectWorkbook.encryptWithPassword",
  "file.info.protectWorkbook.protectCurrentSheet",
  "file.info.protectWorkbook.protectWorkbookStructure",
  "file.new.fromExisting",
  "file.new.templates",
  "file.new.templates.budget",
  "file.new.templates.calendar",
  "file.new.templates.invoice",
  "file.new.templates.more",
  "file.open.pinned",
  "file.open.pinned.kpis",
  "file.open.pinned.q4",
  "file.open.recent",
  "file.open.recent.book1",
  "file.open.recent.budget",
  "file.open.recent.forecast",
  "file.open.recent.more",
  "file.options.account",
  "file.options.options",
  "file.share.email",
  "file.share.email.attachment",
  "file.share.email.link",
  "file.share.presentOnline",
  "file.share.share",
  "formulas.calculation.calculateNow",
  "formulas.calculation.calculateSheet",
  "formulas.calculation.calculationOptions",
  "formulas.definedNames.createFromSelection",
  "formulas.definedNames.defineName",
  "formulas.definedNames.nameManager",
  "formulas.definedNames.useInFormula",
  "formulas.formulaAuditing.errorChecking",
  "formulas.formulaAuditing.evaluateFormula",
  "formulas.formulaAuditing.watchWindow",
  "formulas.functionLibrary.autoSum",
  "formulas.functionLibrary.autoSum.average",
  "formulas.functionLibrary.autoSum.countNumbers",
  "formulas.functionLibrary.autoSum.max",
  "formulas.functionLibrary.autoSum.min",
  "formulas.functionLibrary.autoSum.moreFunctions",
  "formulas.functionLibrary.autoSum.sum",
  "formulas.functionLibrary.dateTime",
  "formulas.functionLibrary.financial",
  "formulas.functionLibrary.insertFunction",
  "formulas.functionLibrary.logical",
  "formulas.functionLibrary.lookupReference",
  "formulas.functionLibrary.mathTrig",
  "formulas.functionLibrary.moreFunctions",
  "formulas.functionLibrary.recentlyUsed",
  "formulas.functionLibrary.text",
  "formulas.solutions.analysisToolPak",
  "help.support.contactSupport",
  "help.support.feedback",
  "help.support.help",
  "help.support.training",
  "home.alignment.mergeCenter",
  "home.alignment.orientation",
  "home.cells.delete",
  "home.cells.insert",
  "home.clipboard.clipboardPane",
  "home.clipboard.clipboardPane.clearAll",
  "home.clipboard.clipboardPane.open",
  "home.clipboard.clipboardPane.options",
  "home.editing.autoSum.moreFunctions",
  "home.editing.clear",
  "home.editing.clear.clearComments",
  "home.editing.clear.clearHyperlinks",
  "home.editing.fill",
  "home.editing.findSelect",
  "home.editing.sortFilter",
  "home.font.borders",
  "home.font.clearFormatting",
  "home.font.fillColor",
  "home.font.fontColor",
  "home.font.fontName",
  "home.font.fontSize",
  "home.number.moreFormats",
  "home.number.numberFormat",
  "home.styles.cellStyles",
  "home.styles.conditionalFormatting",
  "home.styles.conditionalFormatting.clearRules",
  "home.styles.conditionalFormatting.colorScales",
  "home.styles.conditionalFormatting.dataBars",
  "home.styles.conditionalFormatting.highlightCellsRules",
  "home.styles.conditionalFormatting.iconSets",
  "home.styles.conditionalFormatting.manageRules",
  "home.styles.conditionalFormatting.topBottomRules",
  "insert.addins.getAddins",
  "insert.addins.myAddins",
  "insert.charts.area",
  "insert.charts.area.area",
  "insert.charts.area.more",
  "insert.charts.area.stackedArea",
  "insert.charts.bar",
  "insert.charts.bar.clusteredBar",
  "insert.charts.bar.more",
  "insert.charts.bar.stackedBar",
  "insert.charts.boxWhisker",
  "insert.charts.column",
  "insert.charts.column.clusteredColumn",
  "insert.charts.column.more",
  "insert.charts.column.stackedColumn",
  "insert.charts.column.stackedColumn100",
  "insert.charts.combo",
  "insert.charts.funnel",
  "insert.charts.histogram",
  "insert.charts.line",
  "insert.charts.line.line",
  "insert.charts.line.lineWithMarkers",
  "insert.charts.line.more",
  "insert.charts.line.stackedArea",
  "insert.charts.map",
  "insert.charts.map.filledMap",
  "insert.charts.map.more",
  "insert.charts.pie",
  "insert.charts.pie.doughnut",
  "insert.charts.pie.more",
  "insert.charts.pie.pie",
  "insert.charts.pivotChart",
  "insert.charts.radar",
  "insert.charts.recommendedCharts",
  "insert.charts.recommendedCharts.column",
  "insert.charts.recommendedCharts.line",
  "insert.charts.recommendedCharts.more",
  "insert.charts.recommendedCharts.pie",
  "insert.charts.scatter",
  "insert.charts.scatter.more",
  "insert.charts.scatter.scatter",
  "insert.charts.scatter.smoothLines",
  "insert.charts.stock",
  "insert.charts.sunburst",
  "insert.charts.surface",
  "insert.charts.treemap",
  "insert.charts.waterfall",
  "insert.comments.comment",
  "insert.comments.note",
  "insert.equations.equation",
  "insert.equations.inkEquation",
  "insert.filters.slicer",
  "insert.filters.slicer.reportConnections",
  "insert.filters.timeline",
  "insert.illustrations.icons",
  "insert.illustrations.screenshot",
  "insert.illustrations.shapes",
  "insert.illustrations.shapes.arrows",
  "insert.illustrations.shapes.basicShapes",
  "insert.illustrations.shapes.callouts",
  "insert.illustrations.shapes.flowchart",
  "insert.illustrations.shapes.lines",
  "insert.illustrations.shapes.rectangles",
  "insert.illustrations.smartArt",
  "insert.links.link",
  "insert.pivotcharts.pivotChart",
  "insert.pivotcharts.recommendedPivotCharts",
  "insert.sparklines.column",
  "insert.sparklines.line",
  "insert.sparklines.winLoss",
  "insert.symbols.equation",
  "insert.symbols.symbol",
  "insert.tables.pivotTable.fromDataModel",
  "insert.tables.pivotTable.fromExternal",
  "insert.tables.recommendedPivotTables",
  "insert.tables.table",
  "insert.text.headerFooter",
  "insert.text.object",
  "insert.text.signatureLine",
  "insert.text.textBox",
  "insert.text.wordArt",
  "insert.tours.3dMap",
  "insert.tours.launchTour",
  "pageLayout.arrange.align",
  "pageLayout.arrange.align.alignBottom",
  "pageLayout.arrange.align.alignCenter",
  "pageLayout.arrange.align.alignLeft",
  "pageLayout.arrange.align.alignMiddle",
  "pageLayout.arrange.align.alignRight",
  "pageLayout.arrange.align.alignTop",
  "pageLayout.arrange.group",
  "pageLayout.arrange.group.group",
  "pageLayout.arrange.group.regroup",
  "pageLayout.arrange.group.ungroup",
  "pageLayout.arrange.rotate",
  "pageLayout.arrange.rotate.flipHorizontal",
  "pageLayout.arrange.rotate.flipVertical",
  "pageLayout.arrange.rotate.rotateLeft90",
  "pageLayout.arrange.rotate.rotateRight90",
  "pageLayout.arrange.selectionPane",
  "pageLayout.pageSetup.background",
  "pageLayout.pageSetup.background.background",
  "pageLayout.pageSetup.background.delete",
  "pageLayout.pageSetup.breaks",
  "pageLayout.pageSetup.breaks.insertPageBreak",
  "pageLayout.pageSetup.breaks.removePageBreak",
  "pageLayout.pageSetup.breaks.resetAll",
  "pageLayout.pageSetup.margins",
  "pageLayout.pageSetup.orientation",
  "pageLayout.pageSetup.printArea",
  "pageLayout.pageSetup.printTitles",
  "pageLayout.pageSetup.printTitles.printTitles",
  "pageLayout.pageSetup.size",
  "pageLayout.scaleToFit.height",
  "pageLayout.scaleToFit.height.1page",
  "pageLayout.scaleToFit.height.2pages",
  "pageLayout.scaleToFit.height.automatic",
  "pageLayout.scaleToFit.scale",
  "pageLayout.scaleToFit.scale.100",
  "pageLayout.scaleToFit.scale.70",
  "pageLayout.scaleToFit.scale.80",
  "pageLayout.scaleToFit.scale.90",
  "pageLayout.scaleToFit.scale.more",
  "pageLayout.scaleToFit.width",
  "pageLayout.scaleToFit.width.1page",
  "pageLayout.scaleToFit.width.2pages",
  "pageLayout.scaleToFit.width.automatic",
  "pageLayout.sheetOptions.gridlinesPrint",
  "pageLayout.sheetOptions.gridlinesView",
  "pageLayout.sheetOptions.headingsPrint",
  "pageLayout.sheetOptions.headingsView",
  "pageLayout.themes.colors",
  "pageLayout.themes.colors.colorful",
  "pageLayout.themes.colors.customize",
  "pageLayout.themes.colors.office",
  "pageLayout.themes.effects",
  "pageLayout.themes.effects.intense",
  "pageLayout.themes.effects.moderate",
  "pageLayout.themes.effects.subtle",
  "pageLayout.themes.fonts",
  "pageLayout.themes.fonts.aptos",
  "pageLayout.themes.fonts.customize",
  "pageLayout.themes.fonts.office",
  "pageLayout.themes.themes",
  "pageLayout.themes.themes.customize",
  "pageLayout.themes.themes.facet",
  "pageLayout.themes.themes.integral",
  "pageLayout.themes.themes.office",
  "review.changes.protectShareWorkbook",
  "review.changes.protectShareWorkbook.protectWorkbook",
  "review.changes.shareWorkbook",
  "review.changes.shareWorkbook.shareNow",
  "review.changes.trackChanges",
  "review.changes.trackChanges.highlight",
  "review.comments.deleteComment",
  "review.comments.deleteComment.deleteAll",
  "review.comments.deleteComment.deleteThread",
  "review.comments.next",
  "review.comments.previous",
  "review.ink.startInking",
  "review.language.language",
  "review.language.language.setProofing",
  "review.language.language.translate",
  "review.language.translate",
  "review.language.translate.translateSelection",
  "review.language.translate.translateSheet",
  "review.notes.editNote",
  "review.notes.newNote",
  "review.notes.showAllNotes",
  "review.notes.showHideNote",
  "review.proofing.accessibility",
  "review.proofing.smartLookup",
  "review.proofing.spelling",
  "review.proofing.spelling.thesaurus",
  "review.proofing.spelling.wordCount",
  "review.protect.allowEditRanges",
  "review.protect.allowEditRanges.new",
  "review.protect.protectSheet",
  "review.protect.protectWorkbook",
  "review.protect.unprotectSheet",
  "review.protect.unprotectWorkbook",
  "view.show.formulaBar",
  "view.show.gridlines",
  "view.show.headings",
  "view.show.ruler",
  "view.window.arrangeAll",
  "view.window.arrangeAll.cascade",
  "view.window.arrangeAll.horizontal",
  "view.window.arrangeAll.tiled",
  "view.window.arrangeAll.vertical",
  "view.window.freezePanes",
  "view.window.hide",
  "view.window.newWindow",
  "view.window.newWindow.newWindowForActiveSheet",
  "view.window.resetWindowPosition",
  "view.window.switchWindows",
  "view.window.switchWindows.window1",
  "view.window.switchWindows.window2",
  "view.window.synchronousScrolling",
  "view.window.unhide",
  "view.window.viewSideBySide",
  "view.workbookViews.customViews",
  "view.workbookViews.customViews.manage",
  "view.workbookViews.normal",
  "view.workbookViews.pageBreakPreview",
  "view.workbookViews.pageLayout",
];

type OverrideKey = "commandOverrides" | "toggleOverrides";

function extractObjectLiteral(source: string, key: OverrideKey): string | null {
  const idx = source.indexOf(`${key}:`);
  if (idx === -1) return null;
  const braceStart = source.indexOf("{", idx);
  if (braceStart === -1) return null;

  let depth = 0;
  let inString: '"' | "'" | "`" | null = null;
  let inLineComment = false;
  let inBlockComment = false;

  for (let i = braceStart; i < source.length; i += 1) {
    const ch = source[i];
    const next = source[i + 1];

    if (inLineComment) {
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (ch === "*" && next === "/") {
        inBlockComment = false;
        i += 1;
      }
      continue;
    }

    if (inString) {
      if (ch === "\\") {
        i += 1;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }

    if (ch === "/" && next === "/") {
      inLineComment = true;
      i += 1;
      continue;
    }

    if (ch === "/" && next === "*") {
      inBlockComment = true;
      i += 1;
      continue;
    }

    if (ch === '"' || ch === "'" || ch === "`") {
      inString = ch;
      continue;
    }

    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(braceStart, i + 1);
    }
  }

  return null;
}

function extractTopLevelStringKeys(objectText: string): string[] {
  const keys: string[] = [];
  let depth = 0;
  let inLineComment = false;
  let inBlockComment = false;

  const skipWhitespace = (idx: number): number => {
    while (idx < objectText.length && /\s/.test(objectText[idx])) idx += 1;
    return idx;
  };

  for (let i = 0; i < objectText.length; i += 1) {
    const ch = objectText[i];
    const next = objectText[i + 1];

    if (inLineComment) {
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (ch === "*" && next === "/") {
        inBlockComment = false;
        i += 1;
      }
      continue;
    }

    if (ch === "/" && next === "/") {
      inLineComment = true;
      i += 1;
      continue;
    }

    if (ch === "/" && next === "*") {
      inBlockComment = true;
      i += 1;
      continue;
    }

    if (ch === "{") {
      depth += 1;
      continue;
    }

    if (ch === "}") {
      depth -= 1;
      continue;
    }

    if (depth !== 1) continue;
    if (ch !== '"' && ch !== "'") continue;

    const quote = ch;
    let j = i + 1;
    let value = "";

    for (; j < objectText.length; j += 1) {
      const c = objectText[j];
      if (c === "\\") {
        value += objectText[j + 1] ?? "";
        j += 1;
        continue;
      }
      if (c === quote) break;
      value += c;
    }

    if (j >= objectText.length) break;

    const k = skipWhitespace(j + 1);
    if (objectText[k] === ":") {
      keys.push(value);
    }

    i = j;
  }

  return keys;
}

function computeImplementedSchemaCommandIds(schemaCommandIdSet: Set<string>): string[] {
  const commandRegistry = new CommandRegistry();

  // `registerDesktopCommands` and other command catalogs depend on a non-null LayoutController to
  // register many canonical ribbon command ids.
  const layoutController = {
    layout: createDefaultLayout({ primarySheetId: "Sheet1" }),
    openPanel: () => {},
    closePanel: () => {},
    setSplitDirection: () => {},
  } as any;

  registerDesktopCommands({
    commandRegistry,
    app: {} as any,
    layoutController,
    themeController: { setThemePreference: () => {} } as any,
    refreshRibbonUiState: () => {},
    applyFormattingToSelection: () => {},
    getActiveCellNumberFormat: () => null,
    getActiveCellIndentLevel: () => 0,
    openFormatCells: () => {},
    showQuickPick: async () => null,
    pageLayoutHandlers: {
      openPageSetupDialog: () => {},
      updatePageSetup: () => {},
      setPrintArea: () => {},
      clearPrintArea: () => {},
      addToPrintArea: () => {},
      exportPdf: () => {},
    },
    findReplace: { openFind: () => {}, openReplace: () => {}, openGoTo: () => {} },
    workbenchFileHandlers: {
      newWorkbook: () => {},
      openWorkbook: () => {},
      saveWorkbook: () => {},
      saveWorkbookAs: () => {},
      setAutoSaveEnabled: () => {},
      print: () => {},
      printPreview: () => {},
      closeWorkbook: () => {},
      quit: () => {},
    },
    openCommandPalette: () => {},
  });

  registerRibbonMacroCommands({
    commandRegistry,
    handlers: {
      openPanel: () => {},
      focusScriptEditorPanel: () => {},
      focusVbaMigratePanel: () => {},
      setPendingMacrosPanelFocus: () => {},
      startMacroRecorder: () => {},
      stopMacroRecorder: () => {},
      isTauri: () => false,
    },
  });

  registerFormatPainterCommand({
    commandRegistry,
    isArmed: () => false,
    arm: () => {},
    disarm: () => {},
  });

  registerDataQueriesCommands({
    commandRegistry,
    layoutController,
    getPowerQueryService: () => null,
    showToast: () => {},
    notify: async () => {},
    refreshRibbonUiState: () => {},
    focusAfterExecute: () => {},
  });

  const implemented = new Set<string>();
  for (const id of schemaCommandIdSet) {
    if (commandRegistry.getCommand(id)) implemented.add(id);
  }

  const mainTsPath = fileURLToPath(new URL("../../main.ts", import.meta.url));
  const source = readFileSync(mainTsPath, "utf8");

  const addIfSchema = (id: string) => {
    if (schemaCommandIdSet.has(id)) implemented.add(id);
  };

  for (const match of source.matchAll(/case\s+["']([^"']+)["']/g)) {
    addIfSchema(match[1]!);
  }

  for (const match of source.matchAll(/commandId\s*===\s*["']([^"']+)["']/g)) {
    addIfSchema(match[1]!);
  }

  for (const key of ["commandOverrides", "toggleOverrides"] as const satisfies OverrideKey[]) {
    const obj = extractObjectLiteral(source, key);
    if (!obj) continue;
    for (const overrideId of extractTopLevelStringKeys(obj)) {
      addIfSchema(overrideId);
    }
  }

  // Prefix handlers in `handleRibbonCommand`.
  const cellStylesPrefix = "home.styles.cellStyles.";
  if (source.includes(cellStylesPrefix)) {
    for (const id of schemaCommandIdSet) {
      if (id.startsWith(cellStylesPrefix)) implemented.add(id);
    }
  }

  const formatAsTablePrefix = "home.styles.formatAsTable.";
  if (source.includes(formatAsTablePrefix)) {
    for (const id of schemaCommandIdSet) {
      if (!id.startsWith(formatAsTablePrefix)) continue;
      const presetId = id.slice(formatAsTablePrefix.length);
      if (presetId === "light" || presetId === "medium" || presetId === "dark") {
        implemented.add(id);
      }
    }
  }

  return Array.from(implemented).sort();
}

describe("Ribbon command wiring coverage", () => {
  it("classifies every command id in defaultRibbonSchema as implemented vs intentionally unimplemented", () => {
    const schemaCommandIds = collectRibbonCommandIds(defaultRibbonSchema);

    // Guard against a broken traversal so the test can't pass vacuously.
    expect(schemaCommandIds).toContain("clipboard.copy");
    expect(schemaCommandIds).toContain("format.toggleBold");
    expect(schemaCommandIds).toContain("file.open.open");
    expect(schemaCommandIds).toContain("view.zoom.zoom100");

    const implemented = new Set(implementedCommandIds);
    const knownUnimplemented = new Set(knownUnimplementedCommandIds);

    // Sanity checks on the lists themselves.
    const overlap = implementedCommandIds.filter((id) => knownUnimplemented.has(id));
    expect(overlap, `Command ids cannot be both implemented and known-unimplemented: ${overlap.join(", ")}`).toEqual(
      [],
    );

    const unknown = schemaCommandIds.filter((id) => !implemented.has(id) && !knownUnimplemented.has(id));
    expect(
      unknown,
      `Found schema command ids with no wiring classification (add to implementedCommandIds or knownUnimplementedCommandIds):\n${unknown
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);

    const extraImplemented = implementedCommandIds.filter((id) => !schemaCommandIds.includes(id));
    expect(
      extraImplemented,
      `implementedCommandIds contains ids that are not in defaultRibbonSchema:\n${extraImplemented
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);

    const extraKnownUnimplemented = knownUnimplementedCommandIds.filter((id) => !schemaCommandIds.includes(id));
    expect(
      extraKnownUnimplemented,
      `knownUnimplementedCommandIds contains ids that are not in defaultRibbonSchema:\n${extraKnownUnimplemented
        .map((id) => `- ${id}`)
        .join("\n")}`,
    ).toEqual([]);
  });

  it("keeps implementedCommandIds in sync with `apps/desktop/src/main.ts` ribbon wiring", () => {
    const schemaCommandIdSet = new Set(collectRibbonCommandIds(defaultRibbonSchema));
    const expected = computeImplementedSchemaCommandIds(schemaCommandIdSet);
    expect(implementedCommandIds.slice().sort()).toEqual(expected);

    // Keep the denylist honest: it should always represent the remainder of schema ids.
    const expectedSet = new Set(expected);
    const remainder = Array.from(schemaCommandIdSet)
      .filter((id) => !expectedSet.has(id))
      .sort();
    expect(knownUnimplementedCommandIds.slice().sort()).toEqual(remainder);
  });
});
