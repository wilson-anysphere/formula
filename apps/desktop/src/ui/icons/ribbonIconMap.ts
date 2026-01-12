import type { ComponentType } from "react";

import type { IconProps } from "./Icon";
import { AlignBottomIcon } from "./AlignBottomIcon";
import { AlignCenterIcon } from "./AlignCenterIcon";
import { AlignLeftIcon } from "./AlignLeftIcon";
import { AlignMiddleIcon } from "./AlignMiddleIcon";
import { AlignRightIcon } from "./AlignRightIcon";
import { AlignTopIcon } from "./AlignTopIcon";
import { ArrowDownIcon } from "./ArrowDownIcon";
import { ArrowLeftIcon } from "./ArrowLeftIcon";
import { ArrowRightIcon } from "./ArrowRightIcon";
import { ArrowUpIcon } from "./ArrowUpIcon";
import { AutoSumIcon } from "./AutoSumIcon";
import { BoldIcon } from "./BoldIcon";
import { BordersIcon } from "./BordersIcon";
import { BookIcon } from "./BookIcon";
import { BringForwardIcon } from "./BringForwardIcon";
import { CalculatorIcon } from "./CalculatorIcon";
import { CalendarIcon } from "./CalendarIcon";
import { CellStylesIcon } from "./CellStylesIcon";
import { CheckIcon } from "./CheckIcon";
import { ClearFormattingIcon } from "./ClearFormattingIcon";
import { ClearIcon } from "./ClearIcon";
import { ClipboardPaneIcon } from "./ClipboardPaneIcon";
import { ClockIcon } from "./ClockIcon";
import { CloudIcon } from "./CloudIcon";
import { CloseIcon } from "./CloseIcon";
import { CommentIcon } from "./CommentIcon";
import { ColumnWidthIcon } from "./ColumnWidthIcon";
import { CommaIcon } from "./CommaIcon";
import { ConditionalFormattingIcon } from "./ConditionalFormattingIcon";
import { CopyIcon } from "./CopyIcon";
import { CurrencyIcon } from "./CurrencyIcon";
import { CutIcon } from "./CutIcon";
import { DatabaseIcon } from "./DatabaseIcon";
import { DeleteCellsIcon } from "./DeleteCellsIcon";
import { DecreaseDecimalIcon } from "./DecreaseDecimalIcon";
import { DecreaseFontIcon } from "./DecreaseFontIcon";
import { DecreaseIndentIcon } from "./DecreaseIndentIcon";
import { DeleteSheetIcon } from "./DeleteSheetIcon";
import { ExportIcon } from "./ExportIcon";
import { EyeIcon } from "./EyeIcon";
import { EyeOffIcon } from "./EyeOffIcon";
import { FileIcon } from "./FileIcon";
import { ChartIcon } from "./ChartIcon";
import { CodeIcon } from "./CodeIcon";
import { FillColorIcon } from "./FillColorIcon";
import { FillDownIcon } from "./FillDownIcon";
import { FilterIcon } from "./FilterIcon";
import { FindIcon } from "./FindIcon";
import { FeedbackIcon } from "./FeedbackIcon";
import { FolderIcon } from "./FolderIcon";
import { FontColorIcon } from "./FontColorIcon";
import { FontSizeIcon } from "./FontSizeIcon";
import { FunctionIcon } from "./FunctionIcon";
import { FormatAsTableIcon } from "./FormatAsTableIcon";
import { FormatPainterIcon } from "./FormatPainterIcon";
import { GlobeIcon } from "./GlobeIcon";
import { GoToIcon } from "./GoToIcon";
import { GridlinesIcon } from "./GridlinesIcon";
import { HashIcon } from "./HashIcon";
import { HeadingsIcon } from "./HeadingsIcon";
import { IncreaseDecimalIcon } from "./IncreaseDecimalIcon";
import { IncreaseFontIcon } from "./IncreaseFontIcon";
import { IncreaseIndentIcon } from "./IncreaseIndentIcon";
import { InsertCellsIcon } from "./InsertCellsIcon";
import { InsertColumnsIcon } from "./InsertColumnsIcon";
import { InsertRowsIcon } from "./InsertRowsIcon";
import { InsertSheetIcon } from "./InsertSheetIcon";
import { ItalicIcon } from "./ItalicIcon";
import { LayersIcon } from "./LayersIcon";
import { LightningIcon } from "./LightningIcon";
import { LinkIcon } from "./LinkIcon";
import { LockIcon } from "./LockIcon";
import { MailIcon } from "./MailIcon";
import { MacroIcon } from "./MacroIcon";
import { MergeCenterIcon } from "./MergeCenterIcon";
import { MinusIcon } from "./MinusIcon";
import { MoreFormatsIcon } from "./MoreFormatsIcon";
import { MoonIcon } from "./MoonIcon";
import { NoteIcon } from "./NoteIcon";
import { NumberFormatIcon } from "./NumberFormatIcon";
import { OrientationIcon } from "./OrientationIcon";
import { OrganizeSheetsIcon } from "./OrganizeSheetsIcon";
import { PageBreakIcon } from "./PageBreakIcon";
import { PageLandscapeIcon } from "./PageLandscapeIcon";
import { PagePortraitIcon } from "./PagePortraitIcon";
import { PageSetupIcon } from "./PageSetupIcon";
import { PaletteIcon } from "./PaletteIcon";
import { PenIcon } from "./PenIcon";
import { PencilIcon } from "./PencilIcon";
import { PasteIcon } from "./PasteIcon";
import { PasteSpecialIcon } from "./PasteSpecialIcon";
import { PercentIcon } from "./PercentIcon";
import { PiIcon } from "./PiIcon";
import { PinIcon } from "./PinIcon";
import { PictureIcon } from "./PictureIcon";
import { PlayIcon } from "./PlayIcon";
import { PrintIcon } from "./PrintIcon";
import { PrintAreaIcon } from "./PrintAreaIcon";
import { PlusIcon } from "./PlusIcon";
import { PlugIcon } from "./PlugIcon";
import { PuzzleIcon } from "./PuzzleIcon";
import { RedoIcon } from "./RedoIcon";
import { RefreshIcon } from "./RefreshIcon";
import { RecordIcon } from "./RecordIcon";
import { RulerIcon } from "./RulerIcon";
import { SaveIcon } from "./SaveIcon";
import { SendBackwardIcon } from "./SendBackwardIcon";
import { SettingsIcon } from "./SettingsIcon";
import { ShareIcon } from "./ShareIcon";
import { ShieldIcon } from "./ShieldIcon";
import { ReplaceIcon } from "./ReplaceIcon";
import { RowHeightIcon } from "./RowHeightIcon";
import { SlidersIcon } from "./SlidersIcon";
import { SparklesIcon } from "./SparklesIcon";
import { ShapesIcon } from "./ShapesIcon";
import { HighlighterIcon } from "./HighlighterIcon";
import { SideBySideIcon } from "./SideBySideIcon";
import { SmartArtIcon } from "./SmartArtIcon";
import { SnowflakeIcon } from "./SnowflakeIcon";
import { SortFilterIcon } from "./SortFilterIcon";
import { SortIcon } from "./SortIcon";
import { StarIcon } from "./StarIcon";
import { SplitIcon } from "./SplitIcon";
import { StopIcon } from "./StopIcon";
import { StrikethroughIcon } from "./StrikethroughIcon";
import { SubscriptIcon } from "./SubscriptIcon";
import { SuperscriptIcon } from "./SuperscriptIcon";
import { TagIcon } from "./TagIcon";
import { TargetIcon } from "./TargetIcon";
import { TextBoxIcon } from "./TextBoxIcon";
import { TrashIcon } from "./TrashIcon";
import { UnderlineIcon } from "./UnderlineIcon";
import { UndoIcon } from "./UndoIcon";
import { UnlockIcon } from "./UnlockIcon";
import { UserIcon } from "./UserIcon";
import { UsersIcon } from "./UsersIcon";
import { WarningIcon } from "./WarningIcon";
import { WindowIcon } from "./WindowIcon";
import { WrapTextIcon } from "./WrapTextIcon";
import { ZoomInIcon } from "./ZoomInIcon";
import { SunIcon } from "./SunIcon";
import { SyncScrollIcon } from "./SyncScrollIcon";
import { HelpIcon } from "./HelpIcon";
import { GraduationCapIcon } from "./GraduationCapIcon";
import { PhoneIcon } from "./PhoneIcon";

export type RibbonIconComponent = ComponentType<Omit<IconProps, "children">>;

export type RibbonIconId = keyof typeof ribbonIconMap;

/**
 * Look up an icon component by ribbon command id.
 *
 * The ribbon schema currently provides `icon` as a placeholder glyph string.
 * This helper lets the ribbon UI migrate to the internal SVG icon set without
 * coupling the ribbon to a hard-coded import list.
 */
export function getRibbonIcon(commandId: string): RibbonIconComponent | undefined {
  return (ribbonIconMap as Record<string, RibbonIconComponent | undefined>)[commandId];
}

/**
 * Command-id → icon component mapping for ribbon integration.
 *
 * `RibbonButton` will consult this map first (falling back to the schema's
 * placeholder glyph string) so we can progressively replace emoji/text icons
 * with Cursor-style SVGs without rewriting the full ribbon schema.
 */
export const ribbonIconMap = {
  // File
  "file.new.new": FileIcon,
  "file.new.blankWorkbook": FileIcon,
  "file.new.templates": FileIcon,
  "file.info.protectWorkbook": LockIcon,
  "file.info.inspectWorkbook": FindIcon,
  "file.info.manageWorkbook": FolderIcon,
  "file.info.manageWorkbook.versions": ClockIcon,
  "file.open.open": FolderIcon,
  "file.open.recent": ClockIcon,
  "file.open.pinned": PinIcon,
  "file.save.save": SaveIcon,
  "file.save.saveAs": SaveIcon,
  "file.save.autoSave": ClockIcon,
  "file.export.export": ExportIcon,
  "file.export.createPdf": FileIcon,
  "file.export.changeFileType": FileIcon,
  "file.print.print": PrintIcon,
  "file.print.printPreview": EyeIcon,
  "file.print.pageSetup": PageSetupIcon,
  "file.share.share": ShareIcon,
  "file.share.email": MailIcon,
  "file.share.presentOnline": GlobeIcon,
  "file.options.options": SettingsIcon,
  "file.options.account": UserIcon,
  "file.options.close": CloseIcon,

  // Misc generic
  link: LinkIcon,

  // App panels / debug commands (stable ids used by desktop/e2e)
  "open-panel-ai-chat": SparklesIcon,
  "open-ai-panel": SparklesIcon,
  "open-inline-ai-edit": PencilIcon,
  "open-panel-ai-audit": ClipboardPaneIcon,
  "open-ai-audit-panel": ClipboardPaneIcon,
  "open-data-queries-panel": DatabaseIcon,
  "open-macros-panel": MacroIcon,
  "open-script-editor-panel": CodeIcon,
  "open-python-panel": CodeIcon,
  "open-extensions-panel": PuzzleIcon,
  "open-vba-migrate-panel": RefreshIcon,
  "open-comments-panel": CommentIcon,
  "open-marketplace-panel": PuzzleIcon,
  "open-version-history-panel": ClockIcon,
  "open-branch-manager-panel": LayersIcon,

  "audit-precedents": ArrowLeftIcon,
  "audit-dependents": ArrowRightIcon,
  "audit-transitive": RefreshIcon,

  "split-vertical": SplitIcon,
  "split-horizontal": SplitIcon,
  "split-none": CloseIcon,

  "freeze-panes": SnowflakeIcon,
  "freeze-top-row": ArrowUpIcon,
  "freeze-first-column": ArrowLeftIcon,
  "unfreeze-panes": ClearIcon,

  // Clipboard
  "home.clipboard.paste": PasteIcon,
  "home.clipboard.pasteSpecial": PasteSpecialIcon,
  "home.clipboard.cut": CutIcon,
  "home.clipboard.copy": CopyIcon,
  "home.clipboard.formatPainter": FormatPainterIcon,
  "home.clipboard.clipboardPane": ClipboardPaneIcon,

  // Font
  "home.font.fontSize": FontSizeIcon,
  "home.font.increaseFont": IncreaseFontIcon,
  "home.font.decreaseFont": DecreaseFontIcon,
  "home.font.bold": BoldIcon,
  "home.font.italic": ItalicIcon,
  "home.font.underline": UnderlineIcon,
  "home.font.strikethrough": StrikethroughIcon,
  "home.font.subscript": SubscriptIcon,
  "home.font.superscript": SuperscriptIcon,
  "home.font.borders": BordersIcon,
  "home.font.fillColor": FillColorIcon,
  "home.font.fontColor": FontColorIcon,
  "home.font.clearFormatting": ClearFormattingIcon,

  // Alignment
  "home.alignment.topAlign": AlignTopIcon,
  "home.alignment.middleAlign": AlignMiddleIcon,
  "home.alignment.bottomAlign": AlignBottomIcon,
  "home.alignment.alignLeft": AlignLeftIcon,
  "home.alignment.center": AlignCenterIcon,
  "home.alignment.alignRight": AlignRightIcon,
  "home.alignment.orientation": OrientationIcon,
  "home.alignment.wrapText": WrapTextIcon,
  "home.alignment.mergeCenter": MergeCenterIcon,
  "home.alignment.increaseIndent": IncreaseIndentIcon,
  "home.alignment.decreaseIndent": DecreaseIndentIcon,

  // Number
  "home.number.numberFormat": NumberFormatIcon,
  "home.number.accounting": CurrencyIcon,
  "home.number.percent": PercentIcon,
  "home.number.date": CalendarIcon,
  "home.number.comma": CommaIcon,
  "home.number.increaseDecimal": IncreaseDecimalIcon,
  "home.number.decreaseDecimal": DecreaseDecimalIcon,
  "home.number.moreFormats": MoreFormatsIcon,
  "home.number.moreFormats.formatCells": SettingsIcon,
  "home.number.moreFormats.custom": PencilIcon,
  "home.number.formatCells": SettingsIcon,

  // Styles
  "home.styles.conditionalFormatting": ConditionalFormattingIcon,
  "home.styles.formatAsTable": FormatAsTableIcon,
  "home.styles.cellStyles": CellStylesIcon,

  // Cells
  "home.cells.insert": InsertCellsIcon,
  "home.cells.delete": DeleteCellsIcon,
  "home.cells.format": SettingsIcon,
  "home.insert.insertCells": InsertCellsIcon,
  "home.insert.insertRows": InsertRowsIcon,
  "home.insert.insertColumns": InsertColumnsIcon,
  "home.insert.insertSheet": InsertSheetIcon,
  "home.delete.deleteCells": DeleteCellsIcon,
  "home.delete.deleteRows": DeleteCellsIcon,
  "home.delete.deleteColumns": DeleteCellsIcon,
  "home.delete.deleteSheet": DeleteSheetIcon,
  "home.format.formatCells": SettingsIcon,
  "home.format.rowHeight": RowHeightIcon,
  "home.format.columnWidth": ColumnWidthIcon,
  "home.format.organizeSheets": OrganizeSheetsIcon,

  // Editing
  "home.editing.autoSum": AutoSumIcon,
  "home.editing.fill": FillDownIcon,
  "home.editing.clear": ClearIcon,
  "home.editing.sortFilter": SortFilterIcon,

  // Find & Select
  "home.editing.findSelect": FindIcon,
  "home.editing.findSelect.find": FindIcon,
  "home.editing.findSelect.replace": ReplaceIcon,
  "home.editing.findSelect.goTo": GoToIcon,

  // Insert
  "insert.tables.pivotTable": ChartIcon,
  "insert.tables.recommendedPivotTables": StarIcon,
  "insert.tables.table": FormatAsTableIcon,
  "insert.pivotcharts.pivotChart": ChartIcon,
  "insert.pivotcharts.recommendedPivotCharts": StarIcon,
  "insert.illustrations.pictures": PictureIcon,
  "insert.illustrations.onlinePictures": PictureIcon,
  "insert.illustrations.shapes": ShapesIcon,
  "insert.illustrations.icons": StarIcon,
  "insert.illustrations.smartArt": SmartArtIcon,
  "insert.illustrations.screenshot": PictureIcon,
  "insert.addins.getAddins": PlusIcon,
  "insert.addins.myAddins": SmartArtIcon,
  "insert.charts.recommendedCharts": StarIcon,
  "insert.charts.column": ChartIcon,
  "insert.charts.line": ChartIcon,
  "insert.charts.pie": ChartIcon,
  "insert.charts.bar": ChartIcon,
  "insert.charts.area": ChartIcon,
  "insert.charts.scatter": ChartIcon,
  "insert.charts.map": GlobeIcon,
  "insert.charts.histogram": ChartIcon,
  "insert.charts.waterfall": ChartIcon,
  "insert.charts.treemap": ChartIcon,
  "insert.charts.sunburst": ChartIcon,
  "insert.charts.funnel": ChartIcon,
  "insert.charts.boxWhisker": ChartIcon,
  "insert.charts.radar": ChartIcon,
  "insert.charts.surface": ChartIcon,
  "insert.charts.stock": ChartIcon,
  "insert.charts.combo": ChartIcon,
  "insert.charts.pivotChart": ChartIcon,
  "insert.tours.3dMap": GlobeIcon,
  "insert.tours.launchTour": ExportIcon,
  "insert.sparklines.line": ChartIcon,
  "insert.sparklines.column": ChartIcon,
  "insert.sparklines.winLoss": ChartIcon,
  "insert.filters.slicer": FilterIcon,
  "insert.filters.timeline": ClockIcon,
  "insert.links.link": LinkIcon,
  "insert.comments.comment": CommentIcon,
  "insert.comments.note": NoteIcon,
  "insert.text.textBox": TextBoxIcon,
  "insert.text.headerFooter": FileIcon,
  "insert.text.wordArt": FontColorIcon,
  "insert.text.signatureLine": ReplaceIcon,
  "insert.text.object": FileIcon,
  "insert.equations.equation": AutoSumIcon,
  "insert.equations.inkEquation": ReplaceIcon,
  "insert.symbols.equation": AutoSumIcon,
  "insert.symbols.symbol": StarIcon,

  // Page Layout
  "pageLayout.themes.themes": SlidersIcon,
  "pageLayout.themes.colors": PaletteIcon,
  "pageLayout.themes.fonts": FontSizeIcon,
  "pageLayout.themes.effects": SparklesIcon,

  "pageLayout.pageSetup.pageSetupDialog": PageSetupIcon,
  "pageLayout.pageSetup.margins": RulerIcon,
  "pageLayout.pageSetup.orientation": PagePortraitIcon,
  "pageLayout.pageSetup.orientation.portrait": PagePortraitIcon,
  "pageLayout.pageSetup.orientation.landscape": PageLandscapeIcon,
  "pageLayout.pageSetup.size": PagePortraitIcon,
  "pageLayout.pageSetup.printArea": PrintAreaIcon,
  "pageLayout.pageSetup.printArea.set": PrintAreaIcon,
  "pageLayout.pageSetup.printArea.clear": ClearIcon,
  "pageLayout.pageSetup.printArea.addTo": PlusIcon,
  "pageLayout.pageSetup.breaks": PageBreakIcon,
  "pageLayout.pageSetup.background": PictureIcon,
  "pageLayout.pageSetup.printTitles": TagIcon,

  "pageLayout.printArea.setPrintArea": PrintAreaIcon,
  "pageLayout.printArea.clearPrintArea": ClearIcon,

  "pageLayout.export.exportPdf": ExportIcon,

  "pageLayout.scaleToFit.width": ColumnWidthIcon,
  "pageLayout.scaleToFit.height": RowHeightIcon,
  "pageLayout.scaleToFit.scale": PercentIcon,

  "pageLayout.sheetOptions.gridlinesView": GridlinesIcon,
  "pageLayout.sheetOptions.gridlinesPrint": PrintIcon,
  "pageLayout.sheetOptions.headingsView": HeadingsIcon,
  "pageLayout.sheetOptions.headingsPrint": PrintIcon,

  "pageLayout.arrange.bringForward": BringForwardIcon,
  "pageLayout.arrange.sendBackward": SendBackwardIcon,
  "pageLayout.arrange.selectionPane": LayersIcon,
  "pageLayout.arrange.align": AlignCenterIcon,
  "pageLayout.arrange.group": LinkIcon,
  "pageLayout.arrange.group.group": LinkIcon,
  "pageLayout.arrange.group.ungroup": CloseIcon,
  "pageLayout.arrange.group.regroup": RedoIcon,
  "pageLayout.arrange.rotate": RedoIcon,
  "pageLayout.arrange.rotate.rotateRight90": RedoIcon,
  "pageLayout.arrange.rotate.rotateLeft90": UndoIcon,

  // Formulas
  "formulas.functionLibrary.insertFunction": FunctionIcon,
  "formulas.functionLibrary.autoSum": AutoSumIcon,
  "formulas.functionLibrary.recentlyUsed": ClockIcon,
  "formulas.functionLibrary.financial": CurrencyIcon,
  "formulas.functionLibrary.logical": FunctionIcon,
  "formulas.functionLibrary.text": FontSizeIcon,
  "formulas.functionLibrary.dateTime": CalendarIcon,
  "formulas.functionLibrary.lookupReference": FindIcon,
  "formulas.functionLibrary.mathTrig": PiIcon,
  "formulas.functionLibrary.moreFunctions": PlusIcon,

  "formulas.definedNames.nameManager": TagIcon,
  "formulas.definedNames.defineName": PlusIcon,
  "formulas.definedNames.useInFormula": FunctionIcon,
  "formulas.definedNames.createFromSelection": GridlinesIcon,

  "formulas.formulaAuditing.tracePrecedents": ArrowLeftIcon,
  "formulas.formulaAuditing.traceDependents": ArrowRightIcon,
  "formulas.formulaAuditing.removeArrows": CloseIcon,
  "formulas.formulaAuditing.showFormulas": FunctionIcon,
  "formulas.formulaAuditing.errorChecking": WarningIcon,
  "formulas.formulaAuditing.evaluateFormula": CalculatorIcon,
  "formulas.formulaAuditing.watchWindow": EyeIcon,

  "formulas.calculation.calculationOptions": SettingsIcon,
  "formulas.calculation.calculateNow": RefreshIcon,
  "formulas.calculation.calculateSheet": RefreshIcon,

  // Data
  "data.getTransform.getData": FileIcon,
  "data.getTransform.getData.fromFile": FileIcon,
  "data.getTransform.getData.fromDatabase": DatabaseIcon,
  "data.getTransform.getData.fromAzure": CloudIcon,
  "data.getTransform.getData.fromOnlineServices": GlobeIcon,
  "data.getTransform.getData.fromOtherSources": PlusIcon,
  "data.getTransform.recentSources": ClockIcon,
  "data.getTransform.existingConnections": LinkIcon,

  "data.queriesConnections.refreshAll": RefreshIcon,
  "data.queriesConnections.queriesConnections": LayersIcon,
  "data.queriesConnections.properties": SettingsIcon,

  "data.sortFilter.sortAtoZ": SortIcon,
  "data.sortFilter.sortZtoA": SortIcon,
  "data.sortFilter.sort": SortIcon,
  "data.sortFilter.filter": FilterIcon,
  "data.sortFilter.clear": ClearIcon,
  "data.sortFilter.reapply": RefreshIcon,
  "data.sortFilter.advanced": SettingsIcon,

  "data.dataTools.textToColumns": InsertColumnsIcon,
  "data.dataTools.flashFill": LightningIcon,
  "data.dataTools.removeDuplicates": TrashIcon,
  "data.dataTools.dataValidation": CheckIcon,
  "data.dataTools.relationships": LinkIcon,
  "data.dataTools.manageDataModel": SmartArtIcon,

  "data.forecast.whatIfAnalysis": TargetIcon,
  "data.forecast.whatIfAnalysis.scenarioManager": LayersIcon,
  "data.forecast.whatIfAnalysis.goalSeek": TargetIcon,
  "data.forecast.whatIfAnalysis.dataTable": GridlinesIcon,
  "data.forecast.forecastSheet": ChartIcon,

  "data.outline.group": PlusIcon,
  "data.outline.ungroup": MinusIcon,
  "data.outline.subtotal": AutoSumIcon,
  "data.outline.showDetail": PlusIcon,
  "data.outline.hideDetail": MinusIcon,

  "data.dataTypes.stocks": ChartIcon,
  "data.dataTypes.geography": GlobeIcon,

  // Review
  "review.proofing.spelling": CheckIcon,
  "review.proofing.spelling.thesaurus": BookIcon,
  "review.proofing.spelling.wordCount": HashIcon,
  "review.proofing.accessibility": WarningIcon,
  "review.proofing.smartLookup": FindIcon,

  "review.comments.newComment": CommentIcon,
  "review.comments.deleteComment": TrashIcon,
  "review.comments.previous": ArrowUpIcon,
  "review.comments.next": ArrowDownIcon,
  "review.comments.showComments": EyeIcon,

  "review.notes.newNote": NoteIcon,
  "review.notes.editNote": PencilIcon,
  "review.notes.showAllNotes": EyeIcon,
  "review.notes.showHideNote": EyeOffIcon,

  "review.protect.protectSheet": LockIcon,
  "review.protect.unprotectSheet": UnlockIcon,
  "review.protect.protectWorkbook": LockIcon,
  "review.protect.unprotectWorkbook": UnlockIcon,
  "review.protect.allowEditRanges": CheckIcon,
  "review.protect.allowEditRanges.new": PlusIcon,

  "review.ink.startInking": PenIcon,

  "review.language.translate": GlobeIcon,
  "review.language.language": GlobeIcon,

  "review.changes.trackChanges": PencilIcon,
  "review.changes.trackChanges.highlight": HighlighterIcon,
  "review.changes.shareWorkbook": UsersIcon,
  "review.changes.shareWorkbook.shareNow": ShareIcon,
  "review.changes.protectShareWorkbook": LockIcon,

  // View
  "view.workbookViews.normal": GridlinesIcon,
  "view.workbookViews.pageBreakPreview": PageBreakIcon,
  "view.workbookViews.pageLayout": PagePortraitIcon,
  "view.workbookViews.customViews": EyeIcon,

  "view.show.ruler": RulerIcon,
  "view.show.gridlines": GridlinesIcon,
  "view.show.formulaBar": FunctionIcon,
  "view.show.headings": HeadingsIcon,
  "view.show.showFormulas": FunctionIcon,
  "view.show.performanceStats": ChartIcon,

  "view.appearance.theme": PaletteIcon,
  "view.appearance.theme.system": WindowIcon,
  "view.appearance.theme.light": SunIcon,
  "view.appearance.theme.dark": MoonIcon,
  "view.appearance.theme.highContrast": ShieldIcon,

  "view.zoom.zoom": ZoomInIcon,
  "view.zoom.zoom100": PercentIcon,
  "view.zoom.zoomToSelection": TargetIcon,

  "view.window.newWindow": WindowIcon,
  "view.window.arrangeAll": LayersIcon,
  "view.window.freezePanes": SnowflakeIcon,
  "view.window.freezePanes.freezePanes": SnowflakeIcon,
  "view.window.freezePanes.freezeTopRow": ArrowUpIcon,
  "view.window.freezePanes.freezeFirstColumn": ArrowLeftIcon,
  "view.window.freezePanes.unfreeze": ClearIcon,
  "view.window.split": SplitIcon,
  "view.window.hide": EyeOffIcon,
  "view.window.unhide": EyeIcon,
  "view.window.viewSideBySide": SideBySideIcon,
  "view.window.synchronousScrolling": SyncScrollIcon,
  "view.window.resetWindowPosition": UndoIcon,
  "view.window.switchWindows": ReplaceIcon,

  "view.macros.viewMacros": MacroIcon,
  "view.macros.viewMacros.run": PlayIcon,
  "view.macros.viewMacros.edit": PencilIcon,
  "view.macros.viewMacros.delete": TrashIcon,
  "view.macros.recordMacro": RecordIcon,
  "view.macros.recordMacro.stop": StopIcon,
  "view.macros.useRelativeReferences": PinIcon,

  // Developer
  "developer.code.visualBasic": CodeIcon,
  "developer.code.macros": MacroIcon,
  "developer.code.macros.run": PlayIcon,
  "developer.code.macros.edit": PencilIcon,
  "developer.code.recordMacro": RecordIcon,
  "developer.code.recordMacro.stop": StopIcon,
  "developer.code.useRelativeReferences": PinIcon,
  "developer.code.macroSecurity": LockIcon,
  "developer.code.macroSecurity.trustCenter": ShieldIcon,

  "developer.addins.addins": PuzzleIcon,
  "developer.addins.addins.excelAddins": PuzzleIcon,
  "developer.addins.addins.browse": FolderIcon,
  "developer.addins.addins.manage": SettingsIcon,
  "developer.addins.comAddins": PlugIcon,

  "developer.controls.insert": PlusIcon,
  "developer.controls.designMode": SlidersIcon,
  "developer.controls.properties": SettingsIcon,
  "developer.controls.properties.viewProperties": EyeIcon,
  "developer.controls.viewCode": CodeIcon,
  "developer.controls.runDialog": PlayIcon,

  "developer.xml.source": FileIcon,
  "developer.xml.source.refresh": RefreshIcon,
  "developer.xml.mapProperties": GlobeIcon,
  "developer.xml.import": ArrowDownIcon,
  "developer.xml.export": ArrowUpIcon,
  "developer.xml.refreshData": RefreshIcon,

  // Help
  "help.support.help": HelpIcon,
  "help.support.training": GraduationCapIcon,
  "help.support.contactSupport": PhoneIcon,
  "help.support.feedback": FeedbackIcon,

  // Fallbacks (re-usable)
  sort: SortIcon,
  filter: FilterIcon,
  find: FindIcon,
} as const satisfies Record<string, RibbonIconComponent>;
