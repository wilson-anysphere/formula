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
import { CellStylesIcon } from "./CellStylesIcon";
import { ClearFormattingIcon } from "./ClearFormattingIcon";
import { ClearIcon } from "./ClearIcon";
import { ClipboardPaneIcon } from "./ClipboardPaneIcon";
import { ClockIcon } from "./ClockIcon";
import { CloseIcon } from "./CloseIcon";
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
import { IncreaseDecimalIcon } from "./IncreaseDecimalIcon";
import { IncreaseFontIcon } from "./IncreaseFontIcon";
import { IncreaseIndentIcon } from "./IncreaseIndentIcon";
import { InsertCellsIcon } from "./InsertCellsIcon";
import { InsertColumnsIcon } from "./InsertColumnsIcon";
import { InsertRowsIcon } from "./InsertRowsIcon";
import { InsertSheetIcon } from "./InsertSheetIcon";
import { ItalicIcon } from "./ItalicIcon";
import { LinkIcon } from "./LinkIcon";
import { LockIcon } from "./LockIcon";
import { MailIcon } from "./MailIcon";
import { MergeCenterIcon } from "./MergeCenterIcon";
import { MoreFormatsIcon } from "./MoreFormatsIcon";
import { NumberFormatIcon } from "./NumberFormatIcon";
import { OrientationIcon } from "./OrientationIcon";
import { OrganizeSheetsIcon } from "./OrganizeSheetsIcon";
import { PageSetupIcon } from "./PageSetupIcon";
import { PasteIcon } from "./PasteIcon";
import { PasteSpecialIcon } from "./PasteSpecialIcon";
import { PercentIcon } from "./PercentIcon";
import { PinIcon } from "./PinIcon";
import { PrintIcon } from "./PrintIcon";
import { SaveIcon } from "./SaveIcon";
import { SettingsIcon } from "./SettingsIcon";
import { ShareIcon } from "./ShareIcon";
import { ReplaceIcon } from "./ReplaceIcon";
import { RowHeightIcon } from "./RowHeightIcon";
import { SortFilterIcon } from "./SortFilterIcon";
import { SortIcon } from "./SortIcon";
import { StrikethroughIcon } from "./StrikethroughIcon";
import { SubscriptIcon } from "./SubscriptIcon";
import { SuperscriptIcon } from "./SuperscriptIcon";
import { UnderlineIcon } from "./UnderlineIcon";
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

  // Fallbacks (re-usable)
  sort: SortIcon,
  filter: FilterIcon,
  find: FindIcon,
} as const satisfies Record<string, RibbonIconComponent>;
