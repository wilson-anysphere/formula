import type { ComponentType } from "react";

import type { IconProps } from "./Icon";
import { AlignBottomIcon } from "./AlignBottomIcon";
import { AlignCenterIcon } from "./AlignCenterIcon";
import { AlignLeftIcon } from "./AlignLeftIcon";
import { AlignMiddleIcon } from "./AlignMiddleIcon";
import { AlignRightIcon } from "./AlignRightIcon";
import { AlignTopIcon } from "./AlignTopIcon";
import { AutoSumIcon } from "./AutoSumIcon";
import { BoldIcon } from "./BoldIcon";
import { BordersIcon } from "./BordersIcon";
import { BringForwardIcon } from "./BringForwardIcon";
import { CellStylesIcon } from "./CellStylesIcon";
import { ClearFormattingIcon } from "./ClearFormattingIcon";
import { ClearIcon } from "./ClearIcon";
import { ClipboardPaneIcon } from "./ClipboardPaneIcon";
import { ClockIcon } from "./ClockIcon";
import { CloseIcon } from "./CloseIcon";
import { CommentIcon } from "./CommentIcon";
import { ColumnWidthIcon } from "./ColumnWidthIcon";
import { CommaIcon } from "./CommaIcon";
import { ConditionalFormattingIcon } from "./ConditionalFormattingIcon";
import { CopyIcon } from "./CopyIcon";
import { CurrencyIcon } from "./CurrencyIcon";
import { CutIcon } from "./CutIcon";
import { DeleteCellsIcon } from "./DeleteCellsIcon";
import { DecreaseDecimalIcon } from "./DecreaseDecimalIcon";
import { DecreaseFontIcon } from "./DecreaseFontIcon";
import { DecreaseIndentIcon } from "./DecreaseIndentIcon";
import { DeleteSheetIcon } from "./DeleteSheetIcon";
import { ExportIcon } from "./ExportIcon";
import { EyeIcon } from "./EyeIcon";
import { FileIcon } from "./FileIcon";
import { ChartIcon } from "./ChartIcon";
import { FillColorIcon } from "./FillColorIcon";
import { FillDownIcon } from "./FillDownIcon";
import { FilterIcon } from "./FilterIcon";
import { FindIcon } from "./FindIcon";
import { FolderIcon } from "./FolderIcon";
import { FontColorIcon } from "./FontColorIcon";
import { FontSizeIcon } from "./FontSizeIcon";
import { FormatAsTableIcon } from "./FormatAsTableIcon";
import { FormatPainterIcon } from "./FormatPainterIcon";
import { GlobeIcon } from "./GlobeIcon";
import { GoToIcon } from "./GoToIcon";
import { GridlinesIcon } from "./GridlinesIcon";
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
import { LinkIcon } from "./LinkIcon";
import { LockIcon } from "./LockIcon";
import { MailIcon } from "./MailIcon";
import { MergeCenterIcon } from "./MergeCenterIcon";
import { MoreFormatsIcon } from "./MoreFormatsIcon";
import { NoteIcon } from "./NoteIcon";
import { NumberFormatIcon } from "./NumberFormatIcon";
import { OrientationIcon } from "./OrientationIcon";
import { OrganizeSheetsIcon } from "./OrganizeSheetsIcon";
import { PageBreakIcon } from "./PageBreakIcon";
import { PageLandscapeIcon } from "./PageLandscapeIcon";
import { PagePortraitIcon } from "./PagePortraitIcon";
import { PageSetupIcon } from "./PageSetupIcon";
import { PaletteIcon } from "./PaletteIcon";
import { PasteIcon } from "./PasteIcon";
import { PasteSpecialIcon } from "./PasteSpecialIcon";
import { PercentIcon } from "./PercentIcon";
import { PinIcon } from "./PinIcon";
import { PictureIcon } from "./PictureIcon";
import { PrintIcon } from "./PrintIcon";
import { PrintAreaIcon } from "./PrintAreaIcon";
import { PlusIcon } from "./PlusIcon";
import { RedoIcon } from "./RedoIcon";
import { RulerIcon } from "./RulerIcon";
import { SaveIcon } from "./SaveIcon";
import { SendBackwardIcon } from "./SendBackwardIcon";
import { SettingsIcon } from "./SettingsIcon";
import { ShareIcon } from "./ShareIcon";
import { ReplaceIcon } from "./ReplaceIcon";
import { RowHeightIcon } from "./RowHeightIcon";
import { SlidersIcon } from "./SlidersIcon";
import { SparklesIcon } from "./SparklesIcon";
import { ShapesIcon } from "./ShapesIcon";
import { SmartArtIcon } from "./SmartArtIcon";
import { SortFilterIcon } from "./SortFilterIcon";
import { SortIcon } from "./SortIcon";
import { StarIcon } from "./StarIcon";
import { StrikethroughIcon } from "./StrikethroughIcon";
import { SubscriptIcon } from "./SubscriptIcon";
import { SuperscriptIcon } from "./SuperscriptIcon";
import { TagIcon } from "./TagIcon";
import { TextBoxIcon } from "./TextBoxIcon";
import { UnderlineIcon } from "./UnderlineIcon";
import { UndoIcon } from "./UndoIcon";
import { UserIcon } from "./UserIcon";
import { WrapTextIcon } from "./WrapTextIcon";

export type RibbonIconComponent = ComponentType<Omit<IconProps, "children">>;

/**
 * Command-id → icon component mapping for ribbon integration.
 *
 * This file is intentionally not wired into the ribbon UI yet; it exists as a
 * central place to import icons by command id when the ribbon migrates away from
 * placeholder glyph strings.
 */
export const ribbonIconMap = {
  // File
  "file.new.new": FileIcon,
  "file.new.blankWorkbook": FileIcon,
  "file.new.templates": FileIcon,
  "file.info.protectWorkbook": LockIcon,
  "file.info.inspectWorkbook": FindIcon,
  "file.info.manageWorkbook": FolderIcon,
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
  "home.number.comma": CommaIcon,
  "home.number.increaseDecimal": IncreaseDecimalIcon,
  "home.number.decreaseDecimal": DecreaseDecimalIcon,
  "home.number.moreFormats": MoreFormatsIcon,

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

  // Fallbacks (re-usable)
  sort: SortIcon,
  filter: FilterIcon,
  find: FindIcon,
} as const satisfies Record<string, RibbonIconComponent>;
