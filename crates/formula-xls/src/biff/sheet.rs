use std::collections::{BTreeMap, HashMap};

use formula_model::{
    autofilter::{
        FilterColumn, FilterCriterion, FilterJoin, FilterValue, SortCondition, SortState,
    },
    CellRef, Hyperlink, HyperlinkTarget, ManualPageBreaks, Orientation, OutlinePr, PageSetup, Range,
    Scaling, SheetPane, SheetProtection, SheetSelection, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};

use super::records;
use super::strings;

// Record ids used by worksheet parsing.
// See [MS-XLS] sections:
// - ROW: 2.4.184
// - COLINFO: 2.4.48
// - DIMENSIONS: 2.4.84
// - AUTOFILTERINFO: 2.4.29
// - FILTERMODE: 2.4.102
// - Cell records: 2.5.14
// - MULRK: 2.4.141
// - MULBLANK: 2.4.140
const RECORD_ROW: u16 = 0x0208;
const RECORD_COLINFO: u16 = 0x007D;
/// DIMENSIONS [MS-XLS 2.4.84]
const RECORD_DIMENSIONS: u16 = 0x0200;
/// AUTOFILTERINFO [MS-XLS 2.4.29]
const RECORD_AUTOFILTERINFO: u16 = 0x009D;
/// FILTERMODE [MS-XLS 2.4.102]
const RECORD_FILTERMODE: u16 = 0x009B;
/// SORT [MS-XLS 2.4.256]
const RECORD_SORT: u16 = 0x0090;
/// AutoFilter12 [MS-XLS 2.4.30] (Future Record Type; BIFF8 only)
///
/// Excel 2007+ can store filter criteria in BIFF8 using Future Record Type (FRT) records.
/// These records begin with an `FrtHeader`. We detect record types via `FrtHeader.rt`, but
/// Excel typically uses `record_id == rt`.
const RECORD_AUTOFILTER12: u16 = 0x087E;
/// Sort12 (Future Record Type; BIFF8 only)
const RECORD_SORT12: u16 = 0x0880;
/// SortData12 (Future Record Type; BIFF8 only)
const RECORD_SORTDATA12: u16 = 0x0881;
const RECORD_WSBOOL: u16 = 0x0081;
/// MERGEDCELLS [MS-XLS 2.4.139]
const RECORD_MERGEDCELLS: u16 = 0x00E5;

const RECORD_FORMULA: u16 = 0x0006;
const RECORD_BLANK: u16 = 0x0201;
const RECORD_NUMBER: u16 = 0x0203;
const RECORD_LABEL_BIFF5: u16 = 0x0204;
const RECORD_BOOLERR: u16 = 0x0205;
const RECORD_RK: u16 = 0x027E;
const RECORD_RSTRING: u16 = 0x00D6;
const RECORD_LABELSST: u16 = 0x00FD;
const RECORD_MULRK: u16 = 0x00BD;
const RECORD_MULBLANK: u16 = 0x00BE;
/// HLINK [MS-XLS 2.4.110]
const RECORD_HLINK: u16 = 0x01B8;

// Sheet protection records (worksheet substream).
// See [MS-XLS] sections:
// - PROTECT: 2.4.203
// - PASSWORD: 2.4.191
// - OBJPROTECT: 2.4.169
// - SCENPROTECT: 2.4.235
const RECORD_PROTECT: u16 = 0x0012;
const RECORD_PASSWORD: u16 = 0x0013;
const RECORD_OBJPROTECT: u16 = 0x0063;
const RECORD_SCENPROTECT: u16 = 0x00DD;
// Enhanced worksheet protection options (BIFF8 Future Record Type (FRT) records).
//
// Excel stores the richer "Protect Sheet" allow-flags (formatting, inserting/deleting, sort,
// AutoFilter, PivotTables, selection) in the FEAT/FEATHEADR family of records.
//
// See [MS-XLS] sections (record names may vary by version):
// - FEATHEADR / FEAT: shared feature records
// - FEATHEADR11 / FEAT11: Excel 11 variants
const RECORD_FEATHEADR: u16 = 0x0867;
const RECORD_FEAT: u16 = 0x0868;
const RECORD_FEATHEADR11: u16 = 0x0870;
const RECORD_FEAT11: u16 = 0x0871;

// View/UX records (worksheet substream).
// - WINDOW2: [MS-XLS 2.4.354]
// - SCL: [MS-XLS 2.4.244]
// - PANE: [MS-XLS 2.4.174]
// - SELECTION: [MS-XLS 2.4.239]
const RECORD_WINDOW2: u16 = 0x023E;
const RECORD_SCL: u16 = 0x00A0;
const RECORD_PANE: u16 = 0x0041;
const RECORD_SELECTION: u16 = 0x001D;

// Print/page setup records (worksheet substream).
// - SETUP: [MS-XLS 2.4.296]
// - LEFTMARGIN/RIGHTMARGIN/TOPMARGIN/BOTTOMMARGIN: [MS-XLS 2.4.128] etc.
const RECORD_SETUP: u16 = 0x00A1;
const RECORD_LEFTMARGIN: u16 = 0x0026;
const RECORD_RIGHTMARGIN: u16 = 0x0027;
const RECORD_TOPMARGIN: u16 = 0x0028;
const RECORD_BOTTOMMARGIN: u16 = 0x0029;

// Manual page breaks (worksheet substream).
// - VERTICALPAGEBREAKS: [MS-XLS 2.4.349]
// - HORIZONTALPAGEBREAKS: [MS-XLS 2.4.115]
const RECORD_VERTICALPAGEBREAKS: u16 = 0x001A;
const RECORD_HORIZONTALPAGEBREAKS: u16 = 0x001B;

const ROW_HEIGHT_TWIPS_MASK: u16 = 0x7FFF;
const ROW_HEIGHT_DEFAULT_FLAG: u16 = 0x8000;
const ROW_OPTION_HIDDEN: u16 = 0x0020;
const ROW_OPTION_OUTLINE_LEVEL_MASK: u16 = 0x0700;
const ROW_OPTION_OUTLINE_LEVEL_SHIFT: u16 = 8;
const ROW_OPTION_COLLAPSED: u16 = 0x1000;

const COLINFO_OPTION_HIDDEN: u16 = 0x0001;
const COLINFO_OPTION_OUTLINE_LEVEL_MASK: u16 = 0x0700;
const COLINFO_OPTION_OUTLINE_LEVEL_SHIFT: u16 = 8;
const COLINFO_OPTION_COLLAPSED: u16 = 0x1000;

/// Cap warnings collected by best-effort worksheet scans so a crafted file cannot allocate an
/// unbounded number of warning strings.
const MAX_WARNINGS_PER_SHEET: usize = 50;
const WARNINGS_SUPPRESSED_MESSAGE: &str = "additional warnings suppressed";

// WSBOOL (0x0081) is a bitfield of worksheet boolean properties.
//
// Note: In BIFF8, the outline-related flags use inverted semantics:
// - When the bit is set, Excel places the summary *above/left* (non-default).
// - When the bit is clear, Excel uses the default "summary below/right" behaviour.
//
// Empirically, Excel-generated `.xls` fixtures use `WSBOOL=0x0C01` with these bits cleared,
// corresponding to the OOXML defaults (`summaryBelow=true`, `summaryRight=true`).
const WSBOOL_OPTION_ROW_SUMMARY_ABOVE: u16 = 0x0008;
const WSBOOL_OPTION_COL_SUMMARY_LEFT: u16 = 0x0010;
const WSBOOL_OPTION_HIDE_OUTLINE_SYMBOLS: u16 = 0x0040;

fn push_warning_bounded(warnings: &mut Vec<String>, warning: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(warning.into());
        return;
    }
    // Add a single terminal warning so callers have a hint that the import was noisy.
    if warnings.len() == MAX_WARNINGS_PER_SHEET {
        warnings.push(WARNINGS_SUPPRESSED_MESSAGE.to_string());
    }
}

#[derive(Debug, Default, Clone, Copy)]
pub(crate) struct SheetRowProperties {
    pub(crate) height: Option<f32>,
    /// Raw hidden bit from BIFF.
    ///
    /// Excel uses this for both user-hidden rows and rows hidden by a collapsed outline group.
    /// Callers should derive `OutlineEntry.hidden.user` vs `hidden.outline` using the outline
    /// metadata (collapsed summary rows + levels).
    pub(crate) hidden: bool,
    pub(crate) outline_level: u8,
    pub(crate) collapsed: bool,
    /// Default cell XF index (`ixfe`) for the row, when present.
    ///
    /// This corresponds to the legacy `.xls` row-level default format: cells in this row inherit
    /// this format unless they have their own XF.
    pub(crate) xf_index: Option<u16>,
}

#[derive(Debug, Default, Clone, Copy)]
pub(crate) struct SheetColProperties {
    pub(crate) width: Option<f32>,
    /// Raw hidden bit from BIFF.
    pub(crate) hidden: bool,
    pub(crate) outline_level: u8,
    pub(crate) collapsed: bool,
    /// Default cell XF index (`ixfe`) for the column, when present.
    pub(crate) xf_index: Option<u16>,
}

// WINDOW2 grbit flags.
const WINDOW2_FLAG_DSP_GRID: u16 = 0x0002;
const WINDOW2_FLAG_DSP_RW_COL: u16 = 0x0004;
const WINDOW2_FLAG_FROZEN: u16 = 0x0008;
const WINDOW2_FLAG_DSP_ZEROS: u16 = 0x0010;
const WINDOW2_FLAG_FROZEN_NO_SPLIT: u16 = 0x0100;

#[derive(Debug, Default)]
pub(crate) struct SheetRowColProperties {
    pub(crate) rows: BTreeMap<u32, SheetRowProperties>,
    pub(crate) cols: BTreeMap<u32, SheetColProperties>,
    pub(crate) outline_pr: OutlinePr,
    /// Worksheet AutoFilter range inferred from BIFF metadata.
    pub(crate) auto_filter_range: Option<Range>,
    /// AutoFilter filter columns parsed from BIFF AutoFilter12 future records (best-effort).
    pub(crate) auto_filter_columns: Vec<FilterColumn>,
    /// Worksheet AutoFilter sort state, if the sheet substream contained a supported `SORT` record
    /// (or other supported sort record).
    pub(crate) sort_state: Option<SortState>,
    /// Whether the worksheet contained a `FILTERMODE` record (indicating filtered rows).
    pub(crate) filter_mode: bool,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct SheetViewState {
    pub(crate) show_grid_lines: Option<bool>,
    pub(crate) show_headings: Option<bool>,
    pub(crate) show_zeros: Option<bool>,
    pub(crate) zoom: Option<f32>,
    pub(crate) pane: Option<SheetPane>,
    pub(crate) selection: Option<SheetSelection>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct BiffSheetProtection {
    pub(crate) protection: SheetProtection,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct BiffSheetPrintSettings {
    pub(crate) page_setup: Option<PageSetup>,
    pub(crate) manual_page_breaks: ManualPageBreaks,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort parse of worksheet protection state.
///
/// This scans the sheet substream for basic protection records:
/// - `PROTECT` → [`SheetProtection::enabled`]
/// - `PASSWORD` → [`SheetProtection::password_hash`]
/// - `OBJPROTECT` → [`SheetProtection::edit_objects`] (best-effort mapping)
/// - `SCENPROTECT` → [`SheetProtection::edit_scenarios`] (best-effort mapping)
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_protection(
    workbook_stream: &[u8],
    start: usize,
) -> Result<BiffSheetProtection, String> {
    let mut out = BiffSheetProtection::default();

    // FEAT* records can legally be split across one or more `CONTINUE` records. Use the logical
    // iterator so we can reassemble those fragments before decoding.
    let allows_continuation = |record_id: u16| {
        matches!(
            record_id,
            RECORD_FEATHEADR | RECORD_FEAT | RECORD_FEATHEADR11 | RECORD_FEAT11
        )
    };
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data.as_ref();
        match record.record_id {
            RECORD_PROTECT => {
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!("truncated PROTECT record at offset {}", record.offset),
                    );
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.protection.enabled = flag != 0;
            }
            RECORD_PASSWORD => {
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!("truncated PASSWORD record at offset {}", record.offset),
                    );
                    continue;
                }
                let hash = u16::from_le_bytes([data[0], data[1]]);
                out.protection.password_hash = (hash != 0).then_some(hash);
            }
            RECORD_OBJPROTECT => {
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!("truncated OBJPROTECT record at offset {}", record.offset),
                    );
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                // Best-effort mapping: BIFF stores "is protected" flags, while the model stores
                // "is allowed" flags. When objects are protected, editing objects is not allowed.
                out.protection.edit_objects = flag == 0;
            }
            RECORD_SCENPROTECT => {
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!("truncated SCENPROTECT record at offset {}", record.offset),
                    );
                    continue;
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                out.protection.edit_scenarios = flag == 0;
            }
            RECORD_FEATHEADR | RECORD_FEATHEADR11 => {
                match parse_biff_feat_hdr_sheet_protection_allow_mask(data, record.record_id) {
                    Ok(Some(mask)) => apply_sheet_protection_allow_mask(&mut out.protection, mask),
                    Ok(None) => {}
                    Err(err) => push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "failed to parse FEATHEADR record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                }
            }
            RECORD_FEAT | RECORD_FEAT11 => {
                match parse_biff_feat_record_sheet_protection_allow_mask(data, record.record_id) {
                    Ok(Some(mask)) => apply_sheet_protection_allow_mask(&mut out.protection, mask),
                    Ok(None) => {}
                    Err(err) => push_warning_bounded(
                        &mut out.warnings,
                        format!("failed to parse FEAT record at offset {}: {err}", record.offset),
                    ),
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

/// Shared feature type used by Excel to store enhanced worksheet protection options.
///
/// This corresponds to the additional checkboxes in Excel's "Protect Sheet" dialog (formatting,
/// inserting/deleting, sorting, AutoFilter, PivotTables, and selection of locked/unlocked cells).
///
/// NOTE: This is based on [MS-XLS] `isf` (shared feature type) values for FEAT/FEATHEADR records.
const FEAT_ISF_SHEET_PROTECTION: u16 = 0x0002;

// Known allow-flag bits within the enhanced protection mask (best-effort).
//
// The BIFF representation stores a bitmask indicating which actions are allowed when sheet
// protection is enabled. These correspond closely to OOXML `sheetProtection` allow attributes.
//
// The exact packing varies by writer/version; we treat this as a best-effort mapping and ignore
// unknown bits.
const ALLOW_SELECT_LOCKED_CELLS: u32 = 1 << 0;
const ALLOW_SELECT_UNLOCKED_CELLS: u32 = 1 << 1;
const ALLOW_FORMAT_CELLS: u32 = 1 << 2;
const ALLOW_FORMAT_COLUMNS: u32 = 1 << 3;
const ALLOW_FORMAT_ROWS: u32 = 1 << 4;
const ALLOW_INSERT_COLUMNS: u32 = 1 << 5;
const ALLOW_INSERT_ROWS: u32 = 1 << 6;
const ALLOW_INSERT_HYPERLINKS: u32 = 1 << 7;
const ALLOW_DELETE_COLUMNS: u32 = 1 << 8;
const ALLOW_DELETE_ROWS: u32 = 1 << 9;
const ALLOW_SORT: u32 = 1 << 10;
const ALLOW_AUTO_FILTER: u32 = 1 << 11;
const ALLOW_PIVOT_TABLES: u32 = 1 << 12;

fn apply_sheet_protection_allow_mask(protection: &mut SheetProtection, mask: u32) {
    // Selection defaults to true in the model. Some BIFF writers omit the selection bits entirely
    // (relying on defaults), so be conservative: only override selection flags when we see any
    // explicit selection bits set in the mask (either in the low bits or, for some variants, in
    // bits 14/15).
    let low_select = mask & (ALLOW_SELECT_LOCKED_CELLS | ALLOW_SELECT_UNLOCKED_CELLS);
    let high_select = mask & ((1 << 14) | (1 << 15));
    if low_select != 0 {
        protection.select_locked_cells = (mask & ALLOW_SELECT_LOCKED_CELLS) != 0;
        protection.select_unlocked_cells = (mask & ALLOW_SELECT_UNLOCKED_CELLS) != 0;
    } else if high_select != 0 {
        protection.select_locked_cells = (mask & (1 << 14)) != 0;
        protection.select_unlocked_cells = (mask & (1 << 15)) != 0;
    }
    protection.format_cells = (mask & ALLOW_FORMAT_CELLS) != 0;
    protection.format_columns = (mask & ALLOW_FORMAT_COLUMNS) != 0;
    protection.format_rows = (mask & ALLOW_FORMAT_ROWS) != 0;
    protection.insert_columns = (mask & ALLOW_INSERT_COLUMNS) != 0;
    protection.insert_rows = (mask & ALLOW_INSERT_ROWS) != 0;
    protection.insert_hyperlinks = (mask & ALLOW_INSERT_HYPERLINKS) != 0;
    protection.delete_columns = (mask & ALLOW_DELETE_COLUMNS) != 0;
    protection.delete_rows = (mask & ALLOW_DELETE_ROWS) != 0;
    protection.sort = (mask & ALLOW_SORT) != 0;
    protection.auto_filter = (mask & ALLOW_AUTO_FILTER) != 0;
    protection.pivot_tables = (mask & ALLOW_PIVOT_TABLES) != 0;
}

fn frt_payload_start(data: &[u8], expected_rt: u16) -> usize {
    // Many BIFF8 "future record type" (FRT) payloads begin with an `FrtHeader` (8 bytes):
    //   rt (u16), grbitFrt (u16), reserved (u32)
    //
    // When present, the `rt` field typically matches the enclosing record id.
    if data.len() < 8 {
        return 0;
    }
    let rt = u16::from_le_bytes([data[0], data[1]]);
    if rt == expected_rt {
        8
    } else {
        0
    }
}

fn parse_biff_feat_record_sheet_protection_allow_mask(
    data: &[u8],
    record_id: u16,
) -> Result<Option<u32>, String> {
    // Best-effort FEAT record parsing.
    //
    // Common layout (BIFF8):
    // - FrtHeader (8 bytes)
    // - isf (u16) shared feature type
    // - reserved (u16)
    // - cbFeatData (u32)
    // - rgbFeatData (cbFeatData bytes)
    let base = frt_payload_start(data, record_id);
    let header_len = base + 2 + 2 + 4;
    if data.len() < header_len {
        return Err(format!("FEAT record too short (len={})", data.len()));
    }

    let isf = u16::from_le_bytes([data[base], data[base + 1]]);
    if isf != FEAT_ISF_SHEET_PROTECTION {
        return Ok(None);
    }

    let cb_feat_data = u32::from_le_bytes([
        data[base + 4],
        data[base + 5],
        data[base + 6],
        data[base + 7],
    ]) as usize;
    let data_start = base + 8;
    let data_end = data_start
        .checked_add(cb_feat_data)
        .ok_or("FEAT.cbFeatData overflow")?;
    if data.len() < data_end {
        return Err(format!(
            "FEAT record too short for cbFeatData={cb_feat_data} (need {data_end} bytes, got {})",
            data.len()
        ));
    }
    let feat_data = &data[data_start..data_end];
    let mask = parse_allow_mask_best_effort(feat_data)
        .ok_or_else(|| "FEAT protection payload missing allow-mask".to_string())?;
    Ok(Some(mask))
}

fn parse_biff_feat_hdr_sheet_protection_allow_mask(
    data: &[u8],
    record_id: u16,
) -> Result<Option<u32>, String> {
    // Best-effort FEATHEADR record parsing.
    //
    // Common layout (BIFF8):
    // - FrtHeader (8 bytes)
    // - isf (u16) shared feature type
    // - reserved (u16)
    // - cbHdrData (u32)
    // - rgbHdrData (cbHdrData bytes)
    //
    // For the sheet-protection feature, Excel stores the allow-mask in the header data.
    let base = frt_payload_start(data, record_id);
    let header_len = base + 2 + 2 + 4;
    if data.len() < header_len {
        return Err(format!("FEATHEADR record too short (len={})", data.len()));
    }

    let isf = u16::from_le_bytes([data[base], data[base + 1]]);
    if isf != FEAT_ISF_SHEET_PROTECTION {
        return Ok(None);
    }

    let cb_hdr_data = u32::from_le_bytes([
        data[base + 4],
        data[base + 5],
        data[base + 6],
        data[base + 7],
    ]) as usize;
    let data_start = base + 8;
    let data_end = data_start
        .checked_add(cb_hdr_data)
        .ok_or("FEATHEADR.cbHdrData overflow")?;
    if data.len() < data_end {
        return Err(format!(
            "FEATHEADR record too short for cbHdrData={cb_hdr_data} (need {data_end} bytes, got {})",
            data.len()
        ));
    }
    let hdr_data = &data[data_start..data_end];
    let mask = parse_allow_mask_best_effort(hdr_data)
        .ok_or_else(|| "FEATHEADR protection payload missing allow-mask".to_string())?;
    Ok(Some(mask))
}

fn parse_allow_mask_best_effort(payload: &[u8]) -> Option<u32> {
    // Excel's enhanced protection records have evolved over time. Some writers store the allow
    // flags as a 16-bit value, while others may store it as a 32-bit bitfield. Parse both.
    if payload.len() >= 4 {
        let v = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
        return Some(v);
    }
    if payload.len() >= 2 {
        let v = u16::from_le_bytes([payload[0], payload[1]]) as u32;
        return Some(v);
    }
    None
}
/// Best-effort parse of worksheet view/UI state (frozen panes, zoom, selection, display flags).
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_view_state(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetViewState, String> {
    let mut out = SheetViewState::default();

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;

    let mut window2_frozen: Option<bool> = None;
    let mut active_pane: Option<u16> = None;
    let mut selections: Vec<(u16, SheetSelection)> = Vec::new();

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data;
        match record.record_id {
            RECORD_WINDOW2 => match parse_window2_flags(data) {
                Ok(window2) => {
                    out.show_grid_lines = Some(window2.show_grid_lines);
                    out.show_headings = Some(window2.show_headings);
                    out.show_zeros = Some(window2.show_zeros);
                    window2_frozen = Some(window2.frozen_panes);
                }
                Err(err) => push_warning_bounded(
                    &mut out.warnings,
                    format!("failed to parse WINDOW2 record: {err}"),
                ),
            },
            RECORD_SCL => match parse_scl_zoom(data) {
                Ok(zoom) => out.zoom = Some(zoom),
                Err(err) => {
                    push_warning_bounded(&mut out.warnings, format!("failed to parse SCL record: {err}"))
                }
            },
            RECORD_PANE => match parse_pane_record(data, window2_frozen) {
                Ok((pane, pnn_act)) => {
                    out.pane = Some(pane);
                    active_pane = Some(pnn_act);
                }
                Err(err) => push_warning_bounded(
                    &mut out.warnings,
                    format!("failed to parse PANE record: {err}"),
                ),
            },
            RECORD_SELECTION => match parse_selection_record_best_effort(data) {
                Ok((pane, selection)) => selections.push((pane, selection)),
                Err(err) => push_warning_bounded(
                    &mut out.warnings,
                    format!("failed to parse SELECTION record: {err}"),
                ),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    // Prefer the SELECTION record for the active pane, if known; otherwise take the first.
    if let Some(sel) = select_active_selection(active_pane, selections) {
        out.selection = Some(sel);
    }

    Ok(out)
}

#[derive(Debug, Clone, Default)]
pub(crate) struct BiffSheetManualPageBreaks {
    pub(crate) manual_page_breaks: ManualPageBreaks,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort parse of worksheet manual page breaks (`HORIZONTALPAGEBREAKS` / `VERTICALPAGEBREAKS`).
///
/// In BIFF8, the `HorzBrk.row`/`VertBrk.col` fields store the **0-based index** of the first row/col
/// *after* the break. The model stores page breaks as the 0-based row/col index **after which** the
/// break occurs; we therefore subtract 1 (saturating) when importing.
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_manual_page_breaks(
    workbook_stream: &[u8],
    start: usize,
) -> Result<BiffSheetManualPageBreaks, String> {
    let mut out = BiffSheetManualPageBreaks::default();

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data;
        match record.record_id {
            RECORD_HORIZONTALPAGEBREAKS => {
                parse_horizontal_page_breaks_record(
                    data,
                    record.offset,
                    &mut out.manual_page_breaks,
                    &mut out.warnings,
                );
            }
            RECORD_VERTICALPAGEBREAKS => {
                parse_vertical_page_breaks_record(
                    data,
                    record.offset,
                    &mut out.manual_page_breaks,
                    &mut out.warnings,
                );
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn parse_horizontal_page_breaks_record(
    data: &[u8],
    record_offset: usize,
    manual_page_breaks: &mut ManualPageBreaks,
    warnings: &mut Vec<String>,
) {
    // HorizontalPageBreaks payload:
    // - cbrk (u16)
    // - HorzBrk[cbrk] (6 bytes each): row (u16), colStart (u16), colEnd (u16)
    if data.len() < 2 {
        push_warning_bounded(
            warnings,
            format!("truncated HorizontalPageBreaks record at offset {record_offset}"),
        );
        return;
    }

    let cbrk = u16::from_le_bytes([data[0], data[1]]) as usize;

    let mut parsed = 0usize;
    let mut pos = 2usize;
    for _ in 0..cbrk {
        let Some(chunk) = data.get(pos..pos + 6) else {
            break;
        };
        pos = pos.saturating_add(6);
        parsed = parsed.saturating_add(1);

        let row = u16::from_le_bytes([chunk[0], chunk[1]]);
        manual_page_breaks
            .row_breaks_after
            .insert(row.saturating_sub(1) as u32);
    }

    if parsed < cbrk {
        push_warning_bounded(
            warnings,
            format!(
                "truncated HorizontalPageBreaks record at offset {record_offset}: expected {cbrk} breaks, got {parsed}"
            ),
        );
    }
}

fn parse_vertical_page_breaks_record(
    data: &[u8],
    record_offset: usize,
    manual_page_breaks: &mut ManualPageBreaks,
    warnings: &mut Vec<String>,
) {
    // VerticalPageBreaks payload:
    // - cbrk (u16)
    // - VertBrk[cbrk] (6 bytes each): col (u16), rowStart (u16), rowEnd (u16)
    if data.len() < 2 {
        push_warning_bounded(
            warnings,
            format!("truncated VerticalPageBreaks record at offset {record_offset}"),
        );
        return;
    }

    let cbrk = u16::from_le_bytes([data[0], data[1]]) as usize;

    let mut parsed = 0usize;
    let mut pos = 2usize;
    for _ in 0..cbrk {
        let Some(chunk) = data.get(pos..pos + 6) else {
            break;
        };
        pos = pos.saturating_add(6);
        parsed = parsed.saturating_add(1);

        let col = u16::from_le_bytes([chunk[0], chunk[1]]);
        manual_page_breaks
            .col_breaks_after
            .insert(col.saturating_sub(1) as u32);
    }

    if parsed < cbrk {
        push_warning_bounded(
            warnings,
            format!(
                "truncated VerticalPageBreaks record at offset {record_offset}: expected {cbrk} breaks, got {parsed}"
            ),
        );
    }
}
/// Best-effort parse of worksheet print/page setup settings (margins, scaling, paper size, etc).
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_print_settings(
    workbook_stream: &[u8],
    start: usize,
) -> Result<BiffSheetPrintSettings, String> {
    let mut out = BiffSheetPrintSettings::default();

    let mut page_setup = PageSetup::default();
    let mut saw_any_record = false;

    // We need to consult WSBOOL.fFitToPage to decide whether SETUP.iScale or SETUP.iFit* apply.
    // Keep the raw SETUP fields around and compute scaling at the end so record order doesn't
    // matter and "last wins" semantics are respected.
    let mut setup_scale: Option<u16> = None;
    let mut setup_fit_width: Option<u16> = None;
    let mut setup_fit_height: Option<u16> = None;
    let mut wsbool_fit_to_page: Option<bool> = None;

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data;
        match record.record_id {
            // Manual page breaks (best-effort; can be empty).
            RECORD_HORIZONTALPAGEBREAKS => parse_horizontal_page_breaks_record(
                data,
                record.offset,
                &mut out.manual_page_breaks,
                &mut out.warnings,
            ),
            RECORD_VERTICALPAGEBREAKS => parse_vertical_page_breaks_record(
                data,
                record.offset,
                &mut out.manual_page_breaks,
                &mut out.warnings,
            ),
            // Page setup/margins/scaling.
            RECORD_SETUP => {
                saw_any_record = true;
                // SETUP [MS-XLS 2.4.296]
                //
                // Payload (BIFF8):
                // - iPaperSize:u16
                // - iScale:u16
                // - iPageStart:u16 (unused)
                // - iFitWidth:u16
                // - iFitHeight:u16
                // - grbit:u16
                // - iRes:u16 (unused)
                // - iVRes:u16 (unused)
                // - numHdr:Xnum (f64)
                // - numFtr:Xnum (f64)
                // - iCopies:u16 (unused)
                if data.len() < 34 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated SETUP record at offset {} (expected 34 bytes, got {})",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }

                let i_paper_size = u16::from_le_bytes([data[0], data[1]]);
                let i_scale = u16::from_le_bytes([data[2], data[3]]);
                let i_fit_width = u16::from_le_bytes([data[6], data[7]]);
                let i_fit_height = u16::from_le_bytes([data[8], data[9]]);
                let grbit = u16::from_le_bytes([data[10], data[11]]);

                let num_hdr = f64::from_le_bytes(data[16..24].try_into().unwrap());
                let num_ftr = f64::from_le_bytes(data[24..32].try_into().unwrap());

                setup_fit_width = Some(i_fit_width);
                setup_fit_height = Some(i_fit_height);

                // grbit flags:
                // - fPortrait (bit1): 0=landscape, 1=portrait
                // - fNoPls (bit2): if set, printer-related fields are undefined and must be ignored
                // - fNoOrient (bit6): if set, fPortrait must be ignored and orientation defaults to portrait
                const GRBIT_F_PORTRAIT: u16 = 0x0002;
                const GRBIT_F_NOPLS: u16 = 0x0004;
                const GRBIT_F_NOORIENT: u16 = 0x0040;

                let f_no_pls = (grbit & GRBIT_F_NOPLS) != 0;
                let f_no_orient = (grbit & GRBIT_F_NOORIENT) != 0;
                let f_portrait = (grbit & GRBIT_F_PORTRAIT) != 0;

                if !f_no_pls {
                    page_setup.paper_size.code = i_paper_size;
                    setup_scale = Some(i_scale);

                    page_setup.orientation = if f_no_orient {
                        Orientation::Portrait
                    } else if f_portrait {
                        Orientation::Portrait
                    } else {
                        Orientation::Landscape
                    };
                }

                if num_hdr.is_finite() {
                    page_setup.margins.header = num_hdr;
                } else {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "invalid header margin in SETUP record at offset {}: {num_hdr}",
                            record.offset
                        ),
                    );
                }

                if num_ftr.is_finite() {
                    page_setup.margins.footer = num_ftr;
                } else {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "invalid footer margin in SETUP record at offset {}: {num_ftr}",
                            record.offset
                        ),
                    );
                }
            }
            RECORD_LEFTMARGIN | RECORD_RIGHTMARGIN | RECORD_TOPMARGIN | RECORD_BOTTOMMARGIN => {
                saw_any_record = true;
                let record_name = match record.record_id {
                    RECORD_LEFTMARGIN => "LEFTMARGIN",
                    RECORD_RIGHTMARGIN => "RIGHTMARGIN",
                    RECORD_TOPMARGIN => "TOPMARGIN",
                    RECORD_BOTTOMMARGIN => "BOTTOMMARGIN",
                    _ => unreachable!("checked in match arm"),
                };
                if data.len() < 8 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated {record_name} record at offset {} (expected 8 bytes, got {})",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                let value = f64::from_le_bytes(data[0..8].try_into().unwrap());
                if !value.is_finite() {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "invalid {record_name} value at offset {}: {value}",
                            record.offset
                        ),
                    );
                    continue;
                }
                match record.record_id {
                    RECORD_LEFTMARGIN => page_setup.margins.left = value,
                    RECORD_RIGHTMARGIN => page_setup.margins.right = value,
                    RECORD_TOPMARGIN => page_setup.margins.top = value,
                    RECORD_BOTTOMMARGIN => page_setup.margins.bottom = value,
                    _ => {}
                }
            }
            RECORD_WSBOOL => {
                // WSBOOL [MS-XLS 2.4.376]
                // fFitToPage: bit8 (mask 0x0100)
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated WSBOOL record at offset {} (expected >=2 bytes, got {})",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                saw_any_record = true;
                let grbit = u16::from_le_bytes([data[0], data[1]]);
                wsbool_fit_to_page = Some((grbit & 0x0100) != 0);
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let fit_to_page = wsbool_fit_to_page.unwrap_or(false);
    if fit_to_page {
        if let (Some(width), Some(height)) = (setup_fit_width, setup_fit_height) {
            page_setup.scaling = Scaling::FitTo { width, height };
        } else {
            // Some `.xls` writers omit or truncate the SETUP record even when fit-to-page is
            // enabled. Preserve the fit-to-page *mode* even when the target dimensions are
            // unavailable.
            page_setup.scaling = Scaling::FitTo {
                width: 0,
                height: 0,
            };
        }
    } else {
        let scale = setup_scale.unwrap_or(100);
        page_setup.scaling = Scaling::Percent(if scale == 0 { 100 } else { scale });
    }

    // Only surface a page setup when it results in a non-default `PageSetup`. Worksheets almost
    // always contain a `WSBOOL` record (for outline flags), but that does not necessarily imply
    // any explicit print/page setup metadata.
    if saw_any_record && page_setup != PageSetup::default() {
        out.page_setup = Some(page_setup);
    }

    Ok(out)
}
#[derive(Debug, Clone, Copy)]
struct Window2Flags {
    show_grid_lines: bool,
    show_headings: bool,
    show_zeros: bool,
    frozen_panes: bool,
}

fn parse_window2_flags(data: &[u8]) -> Result<Window2Flags, String> {
    if data.len() < 2 {
        return Err("WINDOW2 record too short".to_string());
    }
    let grbit = u16::from_le_bytes([data[0], data[1]]);
    Ok(Window2Flags {
        show_grid_lines: (grbit & WINDOW2_FLAG_DSP_GRID) != 0,
        show_headings: (grbit & WINDOW2_FLAG_DSP_RW_COL) != 0,
        show_zeros: (grbit & WINDOW2_FLAG_DSP_ZEROS) != 0,
        frozen_panes: (grbit & WINDOW2_FLAG_FROZEN) != 0
            || (grbit & WINDOW2_FLAG_FROZEN_NO_SPLIT) != 0,
    })
}

fn parse_scl_zoom(data: &[u8]) -> Result<f32, String> {
    if data.len() < 4 {
        return Err("SCL record too short".to_string());
    }
    let num = u16::from_le_bytes([data[0], data[1]]);
    let denom = u16::from_le_bytes([data[2], data[3]]);
    if denom == 0 {
        return Err("SCL denominator is zero".to_string());
    }
    let zoom = num as f32 / denom as f32;
    if !zoom.is_finite() || zoom <= 0.0 {
        return Err(format!("invalid zoom scale {num}/{denom}"));
    }
    Ok(zoom)
}

fn parse_pane_record(
    data: &[u8],
    frozen_from_window2: Option<bool>,
) -> Result<(SheetPane, u16), String> {
    // PANE record payload (BIFF8): [x:u16][y:u16][rwTop:u16][colLeft:u16][pnnAct:u16]
    if data.len() < 10 {
        return Err("PANE record too short".to_string());
    }
    let x = u16::from_le_bytes([data[0], data[1]]);
    let y = u16::from_le_bytes([data[2], data[3]]);
    let rw_top = u16::from_le_bytes([data[4], data[5]]);
    let col_left = u16::from_le_bytes([data[6], data[7]]);
    let pnn_act = u16::from_le_bytes([data[8], data[9]]);

    let guessed_frozen = (x == col_left && y == rw_top) && (x != 0 || y != 0);
    let frozen = frozen_from_window2.unwrap_or(guessed_frozen);

    let mut pane = SheetPane::default();
    if frozen {
        pane.frozen_rows = y as u32;
        pane.frozen_cols = x as u32;
    } else {
        pane.x_split = (x != 0).then_some(x as f32);
        pane.y_split = (y != 0).then_some(y as f32);
    }

    // Top-left cell for the bottom-right pane.
    let rw_top_u32 = rw_top as u32;
    let col_left_u32 = col_left as u32;
    if rw_top_u32 < EXCEL_MAX_ROWS && col_left_u32 < EXCEL_MAX_COLS {
        pane.top_left_cell = Some(CellRef::new(rw_top_u32, col_left_u32));
    }

    Ok((pane, pnn_act))
}

fn select_active_selection(
    active_pane: Option<u16>,
    mut selections: Vec<(u16, SheetSelection)>,
) -> Option<SheetSelection> {
    if selections.is_empty() {
        return None;
    }
    if let Some(active) = active_pane {
        if let Some(idx) = selections.iter().position(|(pane, _)| *pane == active) {
            return Some(selections.swap_remove(idx).1);
        }
    }
    selections.into_iter().next().map(|(_, sel)| sel)
}

fn parse_selection_record_best_effort(data: &[u8]) -> Result<(u16, SheetSelection), String> {
    // Try a small set of plausible BIFF8 layouts.
    let mut last_err: Option<String> = None;

    // Layout A: pnn:u8 (1 byte), no padding, refs are RefU (6 bytes).
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU8NoPadRefU) {
        return Ok(v);
    }
    // Layout B: pnn:u8 (1 byte) + 1 byte padding, refs are RefU (6 bytes).
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU8PadRefU) {
        return Ok(v);
    }
    // Layout C: pnn:u16, refs are Ref8 (8 bytes).
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU16Ref8) {
        return Ok(v);
    }

    last_err.get_or_insert_with(|| "unrecognized SELECTION record layout".to_string());
    Err(last_err.unwrap())
}

#[derive(Debug, Clone, Copy)]
enum SelectionLayout {
    PnnU8NoPadRefU,
    PnnU8PadRefU,
    PnnU16Ref8,
}

fn parse_selection_record(
    data: &[u8],
    layout: SelectionLayout,
) -> Result<(u16, SheetSelection), String> {
    let (pane, rw_active, col_active, cref, refs_start, ref_len) = match layout {
        SelectionLayout::PnnU8NoPadRefU => {
            if data.len() < 9 {
                return Err("SELECTION record too short".to_string());
            }
            let pane = data[0] as u16;
            let rw_active = u16::from_le_bytes([data[1], data[2]]);
            let col_active = u16::from_le_bytes([data[3], data[4]]);
            // irefActive at [5..7] ignored.
            let cref = u16::from_le_bytes([data[7], data[8]]);
            (pane, rw_active, col_active, cref, 9usize, 6usize)
        }
        SelectionLayout::PnnU8PadRefU => {
            if data.len() < 10 {
                return Err("SELECTION record too short".to_string());
            }
            let pane = data[0] as u16;
            let rw_active = u16::from_le_bytes([data[2], data[3]]);
            let col_active = u16::from_le_bytes([data[4], data[5]]);
            let cref = u16::from_le_bytes([data[8], data[9]]);
            (pane, rw_active, col_active, cref, 10usize, 6usize)
        }
        SelectionLayout::PnnU16Ref8 => {
            if data.len() < 10 {
                return Err("SELECTION record too short".to_string());
            }
            let pane = u16::from_le_bytes([data[0], data[1]]);
            let rw_active = u16::from_le_bytes([data[2], data[3]]);
            let col_active = u16::from_le_bytes([data[4], data[5]]);
            let cref = u16::from_le_bytes([data[8], data[9]]);
            (pane, rw_active, col_active, cref, 10usize, 8usize)
        }
    };

    let cref_usize = cref as usize;
    let needed = refs_start
        .checked_add(cref_usize.checked_mul(ref_len).ok_or("cref overflow")?)
        .ok_or("SELECTION refs length overflow")?;
    if data.len() < needed {
        return Err(format!(
            "SELECTION record too short for {cref} refs (need {needed} bytes, got {})",
            data.len()
        ));
    }

    let active_row_u32 = rw_active as u32;
    let active_col_u32 = col_active as u32;
    if active_row_u32 >= EXCEL_MAX_ROWS || active_col_u32 >= EXCEL_MAX_COLS {
        return Err(format!(
            "active cell out of bounds: row={active_row_u32} col={active_col_u32}"
        ));
    }
    let active_cell = CellRef::new(active_row_u32, active_col_u32);

    let mut ranges = Vec::with_capacity(cref_usize);
    let mut off = refs_start;
    for _ in 0..cref_usize {
        let range = match layout {
            SelectionLayout::PnnU16Ref8 => {
                let rw_first = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
                let rw_last = u16::from_le_bytes([data[off + 2], data[off + 3]]) as u32;
                let col_first = u16::from_le_bytes([data[off + 4], data[off + 5]]) as u32;
                let col_last = u16::from_le_bytes([data[off + 6], data[off + 7]]) as u32;
                off += 8;
                make_range(rw_first, rw_last, col_first, col_last)?
            }
            SelectionLayout::PnnU8NoPadRefU | SelectionLayout::PnnU8PadRefU => {
                let rw_first = u16::from_le_bytes([data[off], data[off + 1]]) as u32;
                let rw_last = u16::from_le_bytes([data[off + 2], data[off + 3]]) as u32;
                let col_first = data[off + 4] as u32;
                let col_last = data[off + 5] as u32;
                off += 6;
                make_range(rw_first, rw_last, col_first, col_last)?
            }
        };
        ranges.push(range);
    }

    Ok((pane, SheetSelection::new(active_cell, ranges)))
}

fn make_range(rw_first: u32, rw_last: u32, col_first: u32, col_last: u32) -> Result<Range, String> {
    if rw_first >= EXCEL_MAX_ROWS
        || rw_last >= EXCEL_MAX_ROWS
        || col_first >= EXCEL_MAX_COLS
        || col_last >= EXCEL_MAX_COLS
    {
        return Err(format!(
            "range out of bounds: rows {rw_first}..={rw_last}, cols {col_first}..={col_last}"
        ));
    }
    let start = CellRef::new(rw_first.min(rw_last), col_first.min(col_last));
    let end = CellRef::new(rw_first.max(rw_last), col_first.max(col_last));
    Ok(Range::new(start, end))
}

pub(crate) fn parse_biff_sheet_row_col_properties(
    workbook_stream: &[u8],
    start: usize,
    codepage: u16,
) -> Result<SheetRowColProperties, String> {
    let mut props = SheetRowColProperties::default();

    // Best-effort AutoFilter metadata:
    // - DIMENSIONS gives the sheet bounding box.
    // - AUTOFILTERINFO contains the number of filter columns.
    // - FILTERMODE indicates an active filter state (some rows hidden by filter).
    let mut dimensions: Option<(u32, u32, u32, u32)> = None;
    let mut autofilter_cols: Option<u32> = None;
    let mut saw_autofilter_info = false;

    // Best-effort AutoFilter/SORT future records (AutoFilter12 / Sort12 / SortData12).
    //
    // We store columns in a map while parsing to de-duplicate any repeated records.
    let mut autofilter12_columns: BTreeMap<u32, FilterColumn> = BTreeMap::new();
    let mut saw_autofilter12 = false;

    let mut saw_eof = false;
    let mut warned_colinfo_first_oob = false;
    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;
    while let Some(next) = iter.next() {
        let record = match next {
            Ok(record) => record,
            Err(err) => {
                push_warning_bounded(
                    &mut props.warnings,
                    format!("malformed BIFF record in sheet stream: {err}"),
                );
                break;
            }
        };

        // BOF indicates the start of a new substream; stop before yielding the next BOF in case the
        // worksheet is missing its EOF.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            // DIMENSIONS [MS-XLS 2.4.84]
            RECORD_DIMENSIONS => {
                let data = record.data;
                if data.len() < 14 {
                    continue;
                }
                let first_row = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
                let last_row_plus1 = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
                let first_col = u16::from_le_bytes([data[8], data[9]]) as u32;
                let last_col_plus1 = u16::from_le_bytes([data[10], data[11]]) as u32;
                dimensions = Some((first_row, last_row_plus1, first_col, last_col_plus1));
            }
            // AUTOFILTERINFO [MS-XLS 2.4.29]
            RECORD_AUTOFILTERINFO => {
                let data = record.data;
                saw_autofilter_info = true;
                if data.len() < 2 {
                    continue;
                }
                let cols = u16::from_le_bytes([data[0], data[1]]) as u32;
                autofilter_cols = Some(cols);
            }
            // FILTERMODE [MS-XLS 2.4.102]
            RECORD_FILTERMODE => {
                props.filter_mode = true;
            }
            // SORT [MS-XLS 2.4.256]
            RECORD_SORT => match parse_sort_record_best_effort(record.data) {
                Ok(Some(state)) => {
                    // Prefer the last SORT record in the sheet stream (Excel may emit multiple
                    // records as sort state evolves).
                    props.sort_state = Some(state);
                }
                Ok(None) => {}
                Err(err) => push_warning_bounded(
                    &mut props.warnings,
                    format!("failed to parse SORT record at offset {}: {err}", record.offset),
                ),
            },
            // AutoFilter12 / Sort12 / SortData12 (BIFF8 Future Record Type records).
            //
            // These records start with an `FrtHeader` structure. The record id in the BIFF
            // stream is often the same as `FrtHeader.rt`, but we still key off `rt` for
            // robustness.
            id if id >= 0x0850 && id <= 0x08FF => {
                let Some((rt, frt_payload)) = parse_frt_header(record.data) else {
                    // Not a valid FRT header; ignore silently.
                    continue;
                };

                match rt {
                    RECORD_AUTOFILTER12 => {
                        saw_autofilter12 = true;
                        match decode_autofilter12_record(frt_payload, codepage) {
                            Ok(Some(column)) => {
                                autofilter12_columns
                                    .entry(column.col_id)
                                    .or_insert(column);
                            }
                            Ok(None) => {
                                // Record parsed but contained no recoverable values.
                                if !props
                                    .warnings
                                    .iter()
                                    .any(|w| w == "unsupported AutoFilter12")
                                {
                                    push_warning_bounded(
                                        &mut props.warnings,
                                        "unsupported AutoFilter12".to_string(),
                                    );
                                }
                            }
                            Err(_) => {
                                // Best-effort: preserve nothing but surface a deterministic warning.
                                if !props
                                    .warnings
                                    .iter()
                                    .any(|w| w == "unsupported AutoFilter12")
                                {
                                    push_warning_bounded(
                                        &mut props.warnings,
                                        "unsupported AutoFilter12".to_string(),
                                    );
                                }
                            }
                        }
                    }
                    RECORD_SORT12 => {
                        if !props.warnings.iter().any(|w| w == "unsupported Sort12") {
                            push_warning_bounded(&mut props.warnings, "unsupported Sort12");
                        }
                    }
                    RECORD_SORTDATA12 => {
                        if !props
                            .warnings
                            .iter()
                            .any(|w| w == "unsupported SortData12")
                        {
                            push_warning_bounded(&mut props.warnings, "unsupported SortData12");
                        }
                    }
                    _ => {}
                }
            }
            // ROW [MS-XLS 2.4.184]
            RECORD_ROW => {
                let data = record.data;
                if data.len() < 16 {
                    push_warning_bounded(
                        &mut props.warnings,
                        format!(
                            "malformed ROW record at offset {}: expected >=16 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let height_options = u16::from_le_bytes([data[6], data[7]]);
                let height_twips = height_options & ROW_HEIGHT_TWIPS_MASK;
                let default_height = (height_options & ROW_HEIGHT_DEFAULT_FLAG) != 0;
                let options = u16::from_le_bytes([data[12], data[13]]);
                let hidden = (options & ROW_OPTION_HIDDEN) != 0;
                let outline_level = ((options & ROW_OPTION_OUTLINE_LEVEL_MASK)
                    >> ROW_OPTION_OUTLINE_LEVEL_SHIFT) as u8;
                let collapsed = (options & ROW_OPTION_COLLAPSED) != 0;
                let ixfe = u16::from_le_bytes([data[14], data[15]]);
                let xf_index = (ixfe != 0).then_some(ixfe);

                let height =
                    (!default_height && height_twips > 0).then_some(height_twips as f32 / 20.0);

                if hidden
                    || height.is_some()
                    || outline_level > 0
                    || collapsed
                    || xf_index.is_some()
                {
                    let entry = props.rows.entry(row).or_default();
                    if let Some(height) = height {
                        entry.height = Some(height);
                    }
                    if hidden {
                        entry.hidden = true;
                    }
                    if outline_level > 0 {
                        entry.outline_level = outline_level;
                    }
                    if collapsed {
                        entry.collapsed = true;
                    }
                    if xf_index.is_some() {
                        entry.xf_index = xf_index;
                    }
                }
            }
            // COLINFO [MS-XLS 2.4.48]
            RECORD_COLINFO => {
                let data = record.data;
                if data.len() < 12 {
                    push_warning_bounded(
                        &mut props.warnings,
                        format!(
                            "malformed COLINFO record at offset {}: expected >=12 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                let first_col = u16::from_le_bytes([data[0], data[1]]) as u32;
                let last_col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let width_raw = u16::from_le_bytes([data[4], data[5]]);
                let xf_index = u16::from_le_bytes([data[6], data[7]]);
                let options = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (options & COLINFO_OPTION_HIDDEN) != 0;
                let outline_level = ((options & COLINFO_OPTION_OUTLINE_LEVEL_MASK)
                    >> COLINFO_OPTION_OUTLINE_LEVEL_SHIFT)
                    as u8;
                let collapsed = (options & COLINFO_OPTION_COLLAPSED) != 0;

                let width = (width_raw > 0).then_some(width_raw as f32 / 256.0);
                let xf_index = (xf_index != 0).then_some(xf_index);

                if hidden || width.is_some() || outline_level > 0 || collapsed || xf_index.is_some()
                {
                    if first_col > last_col {
                        push_warning_bounded(
                            &mut props.warnings,
                            format!(
                                "malformed COLINFO record at offset {}: first_col ({first_col}) > last_col ({last_col})",
                                record.offset
                            ),
                        );
                        continue;
                    }

                    // Clamp COLINFO column ranges to the model's Excel bounds to avoid excessive
                    // work/memory usage for corrupt files with huge ranges.
                    if first_col >= EXCEL_MAX_COLS {
                        if !warned_colinfo_first_oob {
                            push_warning_bounded(
                                &mut props.warnings,
                                format!(
                                    "ignoring COLINFO record with out-of-bounds first_col ({first_col}) at offset {}",
                                    record.offset
                                ),
                            );
                            warned_colinfo_first_oob = true;
                        }
                        continue;
                    }

                    let max_col = EXCEL_MAX_COLS.saturating_sub(1);
                    let clamped_last_col = last_col.min(max_col);
                    if clamped_last_col != last_col {
                        push_warning_bounded(
                            &mut props.warnings,
                            format!(
                                "COLINFO column range {first_col}..={last_col} truncated to {first_col}..={clamped_last_col}"
                            ),
                        );
                    }

                    for col in first_col..=clamped_last_col {
                        let entry = props.cols.entry(col).or_default();
                        if let Some(width) = width {
                            entry.width = Some(width);
                        }
                        if hidden {
                            entry.hidden = true;
                        }
                        if outline_level > 0 {
                            entry.outline_level = outline_level;
                        }
                        if collapsed {
                            entry.collapsed = true;
                        }
                        if xf_index.is_some() {
                            entry.xf_index = xf_index;
                        }
                    }
                }
            }
            // WSBOOL [MS-XLS 2.4.376]
            RECORD_WSBOOL => {
                let data = record.data;
                if data.len() < 2 {
                    push_warning_bounded(
                        &mut props.warnings,
                        format!(
                            "malformed WSBOOL record at offset {}: expected >=2 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }
                let options = u16::from_le_bytes([data[0], data[1]]);
                props.outline_pr.summary_below = (options & WSBOOL_OPTION_ROW_SUMMARY_ABOVE) == 0;
                props.outline_pr.summary_right = (options & WSBOOL_OPTION_COL_SUMMARY_LEFT) == 0;
                props.outline_pr.show_outline_symbols =
                    (options & WSBOOL_OPTION_HIDE_OUTLINE_SYMBOLS) == 0;
            }
            // EOF terminates the sheet substream.
            records::RECORD_EOF => {
                saw_eof = true;
                break;
            }
            _ => {}
        }
    }

    if !autofilter12_columns.is_empty() {
        props.auto_filter_columns = autofilter12_columns.into_values().collect();
    } else if saw_autofilter12 && props.auto_filter_columns.is_empty() {
        // If we observed an AutoFilter12 record but couldn't decode it, ensure we still emit at
        // least one warning so callers can understand why filter criteria were dropped.
        //
        // Avoid duplicating the warning if the decode logic already pushed one.
        if !props
            .warnings
            .iter()
            .any(|w| w.contains("unsupported AutoFilter12"))
        {
            push_warning_bounded(&mut props.warnings, "unsupported AutoFilter12".to_string());
        }
    }

    // Infer the AutoFilter range when we have enough information.
    //
    // BIFF stores the AutoFilter state across multiple records:
    // - `AUTOFILTERINFO` indicates the presence of an AutoFilter and the number of columns.
    // - `FILTERMODE` indicates that some rows are currently hidden by a filter.
    // - The sheet's `DIMENSIONS` record gives a reasonable bounding box for the filter range.
    //
    // Note: the canonical AutoFilter range is stored in the `_FilterDatabase` defined name, but we
    // may not have that available. We use this DIMENSIONS-based heuristic as a best-effort
    // approximation so the filter dropdown is preserved.
    if props.auto_filter_range.is_none() && (saw_autofilter_info || props.filter_mode) {
        if let Some((first_row, last_row_plus1, first_col, last_col_plus1)) = dimensions {
            // DIMENSIONS uses "last row/col + 1" semantics.
            if last_row_plus1 > 0 && last_col_plus1 > 0 {
                let mut end_row = last_row_plus1.saturating_sub(1);
                let mut end_col = last_col_plus1.saturating_sub(1);

                if first_row >= EXCEL_MAX_ROWS || first_col >= EXCEL_MAX_COLS {
                    // Ignore out-of-bounds dimensions.
                } else {
                    end_row = end_row.min(EXCEL_MAX_ROWS.saturating_sub(1));
                    end_col = end_col.min(EXCEL_MAX_COLS.saturating_sub(1));

                    if let Some(cols) = autofilter_cols {
                        if cols > 0 {
                            let last_filter_col = first_col.saturating_add(cols.saturating_sub(1));
                            end_col = end_col.min(last_filter_col);
                        }
                    }

                    if end_row >= first_row && end_col >= first_col {
                        props.auto_filter_range = Some(Range::new(
                            CellRef::new(first_row, first_col),
                            CellRef::new(end_row, end_col),
                        ));
                    }
                }
            }
        }
    }

    if !saw_eof {
        // Some `.xls` files in the wild omit worksheet EOF records. Treat this as best-effort and
        // return partial state.
        push_warning_bounded(
            &mut props.warnings,
            "unexpected end of worksheet stream (missing EOF)".to_string(),
        );
    }

    Ok(props)
}

// SORT.grbit flags.
//
// We only currently care about whether the sorted range includes a header row. When present, we
// map the sort key range to match OOXML `<sortCondition ref="...">` semantics: the sort key range
// excludes the header row (Excel sorts apply to the data rows beneath the header).
//
// Note: This is intentionally best-effort. BIFF `SORT` supports a range of options (left-to-right
// sorts, custom lists, case sensitivity, locale-aware comparisons) that are not currently
// representable in `formula_model::SortState`.
const SORT_FLAG_HAS_HEADER: u16 = 0x0001;

/// Best-effort parse of a BIFF8 `SORT` record into a model [`SortState`].
///
/// The BIFF `SORT` record predates OOXML and stores sort settings in a compact binary form. Excel
/// 97-2003 supports up to three sort keys. We map each key to a [`SortCondition`] with a column
/// range spanning the sorted rows.
fn parse_sort_record_best_effort(data: &[u8]) -> Result<Option<SortState>, String> {
    // The canonical BIFF8 `SORT` record layout is 24 bytes:
    // - rwFirst, rwLast, colFirst, colLast (Ref8U): 8 bytes
    // - grbit: 2 bytes
    // - cKey: 2 bytes
    // - rgKey[3]: 3 * u16 (column index)
    // - rgOrder[3]: 3 * u16 (0=ascending, 1=descending)
    //
    // Some producers may emit alternative layouts; this parser is intentionally conservative and
    // only decodes the layout we write in tests today.
    if data.len() < 24 {
        return Err(format!(
            "SORT record too short: expected >=24 bytes, got {}",
            data.len()
        ));
    }

    let rw_first = u16::from_le_bytes([data[0], data[1]]) as u32;
    let rw_last = u16::from_le_bytes([data[2], data[3]]) as u32;
    let col_first = u16::from_le_bytes([data[4], data[5]]) as u32;
    let col_last = u16::from_le_bytes([data[6], data[7]]) as u32;
    let grbit = u16::from_le_bytes([data[8], data[9]]);
    let c_keys = u16::from_le_bytes([data[10], data[11]]) as usize;

    if rw_first >= EXCEL_MAX_ROWS || rw_last >= EXCEL_MAX_ROWS {
        return Err(format!(
            "sorted row range out of bounds: {rw_first}..={rw_last}"
        ));
    }
    if col_first >= EXCEL_MAX_COLS || col_last >= EXCEL_MAX_COLS {
        return Err(format!(
            "sorted column range out of bounds: {col_first}..={col_last}"
        ));
    }

    let has_header = (grbit & SORT_FLAG_HAS_HEADER) != 0;
    let mut start_row = rw_first.min(rw_last);
    let end_row = rw_first.max(rw_last);
    if has_header {
        start_row = start_row.saturating_add(1);
    }
    if start_row > end_row {
        // Range contains only a header row (or is otherwise empty).
        return Ok(None);
    }

    let key_cols = [
        u16::from_le_bytes([data[12], data[13]]) as u32,
        u16::from_le_bytes([data[14], data[15]]) as u32,
        u16::from_le_bytes([data[16], data[17]]) as u32,
    ];
    let orders = [
        u16::from_le_bytes([data[18], data[19]]),
        u16::from_le_bytes([data[20], data[21]]),
        u16::from_le_bytes([data[22], data[23]]),
    ];

    let mut conditions = Vec::new();
    for i in 0..c_keys.min(3) {
        let key_col = key_cols[i];
        if key_col == 0xFFFF {
            continue;
        }
        if key_col < col_first.min(col_last) || key_col > col_first.max(col_last) {
            // Key outside the sorted range; ignore.
            continue;
        }
        if key_col >= EXCEL_MAX_COLS {
            continue;
        }

        let descending = orders[i] != 0;
        conditions.push(SortCondition {
            range: Range::new(
                CellRef::new(start_row, key_col),
                CellRef::new(end_row, key_col),
            ),
            descending,
        });
    }

    if conditions.is_empty() {
        Ok(None)
    } else {
        Ok(Some(SortState { conditions }))
    }
}

/// Parse an `FrtHeader` structure and return `(rt, payload_after_header)`.
///
/// `FrtHeader` is an 8-byte structure: `rt` (u16), `grbitFrt` (u16), and a reserved u32.
fn parse_frt_header(data: &[u8]) -> Option<(u16, &[u8])> {
    if data.len() < 8 {
        return None;
    }
    let rt = u16::from_le_bytes([data[0], data[1]]);
    Some((rt, &data[8..]))
}

/// Best-effort decode of an AutoFilter12 record payload (after `FrtHeader`).
///
/// AutoFilter12 is a BIFF8 Future Record Type (FRT) record used by Excel 2007+ to store
/// newer AutoFilter semantics in `.xls` files.
///
/// The on-disk structure is complex. We currently implement a conservative subset that
/// attempts to recover multi-value filters:
/// - Treat the first u16 as `colId` (matching OOXML `filterColumn/@colId`).
/// - Attempt to read a u16 value count from a few common offsets and then decode that many
///   `XLUnicodeString` values.
///
/// If decoding fails, callers should treat the record as unsupported.
fn decode_autofilter12_record(
    payload: &[u8],
    codepage: u16,
) -> Result<Option<FilterColumn>, String> {
    if payload.len() < 2 {
        return Err("AutoFilter12 payload too short".to_string());
    }
    let col_id = u16::from_le_bytes([payload[0], payload[1]]) as u32;

    // Candidate layouts:
    // - [colId:u16][cVals:u16][vals...]
    // - [colId:u16][flags:u16][cVals:u16][vals...]
    // - [colId:u16][flags:u16][unused:u16][cVals:u16][vals...]
    const MAX_VALUES: usize = 1024;
    let candidates: &[(usize, usize)] = &[(2, 4), (4, 6), (6, 8)];

    for &(count_off, vals_off) in candidates {
        if payload.len() < vals_off {
            continue;
        }
        let count_bytes = payload.get(count_off..count_off + 2).ok_or_else(|| {
            format!(
                "AutoFilter12 payload too short for count at offset {count_off} (len={})",
                payload.len()
            )
        })?;
        let mut count = u16::from_le_bytes([count_bytes[0], count_bytes[1]]) as usize;
        if count == 0 {
            continue;
        }
        // Basic sanity check: a BIFF8 XLUnicodeString is at least 3 bytes (cch:u16 + flags:u8).
        if count > MAX_VALUES || count.saturating_mul(3) > payload.len().saturating_sub(vals_off) {
            continue;
        }

        let mut pos = vals_off;
        let mut values: Vec<String> = Vec::with_capacity(count.min(16));
        while count > 0 {
            let rest = payload.get(pos..).unwrap_or_default();
            let Ok((mut s, used)) = strings::parse_biff8_unicode_string(rest, codepage) else {
                values.clear();
                break;
            };
            if s.contains('\0') {
                s.retain(|ch| ch != '\0');
            }
            pos = pos.saturating_add(used);
            values.push(s);
            count -= 1;
        }

        if values.is_empty() {
            continue;
        }

        let mut criteria = Vec::with_capacity(values.len());
        for v in &values {
            if let Ok(n) = v.parse::<f64>() {
                criteria.push(FilterCriterion::Equals(FilterValue::Number(n)));
            } else {
                criteria.push(FilterCriterion::Equals(FilterValue::Text(v.clone())));
            }
        }

        return Ok(Some(FilterColumn {
            col_id,
            join: FilterJoin::Any,
            criteria,
            values,
            raw_xml: Vec::new(),
        }));
    }

    Ok(None)
}

/// Parse merged cell regions from a worksheet BIFF substream.
///
/// `calamine` usually exposes merge ranges via `worksheet_merge_cells()`, but some `.xls` files in
/// the wild contain `MERGEDCELLS` records that are not surfaced (or surfaced incompletely). This is
/// a best-effort fallback that scans the sheet substream directly and recovers any merge ranges it
/// can.
pub(crate) fn parse_biff_sheet_merged_cells(
    workbook_stream: &[u8],
    start: usize,
) -> Result<Vec<Range>, String> {
    let mut out = Vec::new();

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        match record.record_id {
            RECORD_MERGEDCELLS => {
                // MERGEDCELLS [MS-XLS 2.4.139]
                // - cAreas (2 bytes): number of Ref8 structures
                // - Ref8 (8 bytes each): rwFirst, rwLast, colFirst, colLast (all u16)
                let data = record.data;
                if data.len() < 2 {
                    continue;
                }

                let c_areas = u16::from_le_bytes([data[0], data[1]]) as usize;
                let mut pos = 2usize;
                for _ in 0..c_areas {
                    let Some(chunk) = data.get(pos..pos + 8) else {
                        break;
                    };
                    pos = pos.saturating_add(8);

                    let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]) as u32;
                    let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]) as u32;
                    let col_first = u16::from_le_bytes([chunk[4], chunk[5]]) as u32;
                    let col_last = u16::from_le_bytes([chunk[6], chunk[7]]) as u32;

                    if rw_first >= EXCEL_MAX_ROWS
                        || rw_last >= EXCEL_MAX_ROWS
                        || col_first >= EXCEL_MAX_COLS
                        || col_last >= EXCEL_MAX_COLS
                    {
                        // Ignore out-of-bounds ranges to avoid corrupt coordinates.
                        continue;
                    }

                    out.push(Range::new(
                        CellRef::new(rw_first, col_first),
                        CellRef::new(rw_last, col_last),
                    ));
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

pub(crate) fn parse_biff_sheet_cell_xf_indices_filtered(
    workbook_stream: &[u8],
    start: usize,
    xf_is_interesting: Option<&[bool]>,
) -> Result<HashMap<CellRef, u16>, String> {
    let mut out = HashMap::new();

    let mut maybe_insert = |row: u32, col: u32, xf: u16| {
        if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
            return;
        }
        if let Some(mask) = xf_is_interesting {
            let idx = xf as usize;
            // Retain out-of-range XF indices so callers can surface an aggregated warning.
            if idx >= mask.len() {
                out.insert(CellRef::new(row, col), xf);
                return;
            }
            if !mask[idx] {
                return;
            }
        }
        out.insert(CellRef::new(row, col), xf);
    };

    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        let data = record.data;
        match record.record_id {
            // Cell records with a `Cell` header (rw, col, ixfe) [MS-XLS 2.5.14].
            //
            // We only care about extracting the XF index (`ixfe`) so we can resolve
            // number formats from workbook globals.
            RECORD_FORMULA | RECORD_BLANK | RECORD_NUMBER | RECORD_LABEL_BIFF5 | RECORD_BOOLERR
            | RECORD_RK | RECORD_RSTRING | RECORD_LABELSST => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let xf = u16::from_le_bytes([data[4], data[5]]);
                maybe_insert(row, col, xf);
            }
            // MULRK [MS-XLS 2.4.141]
            RECORD_MULRK => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last =
                    u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]) as u32;
                let rk_data = &data[4..data.len().saturating_sub(2)];
                for (idx, chunk) in rk_data.chunks_exact(6).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf);
                }
            }
            // MULBLANK [MS-XLS 2.4.140]
            RECORD_MULBLANK => {
                if data.len() < 6 {
                    continue;
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col_first = u16::from_le_bytes([data[2], data[3]]) as u32;
                let col_last =
                    u16::from_le_bytes([data[data.len() - 2], data[data.len() - 1]]) as u32;
                let xf_data = &data[4..data.len().saturating_sub(2)];
                for (idx, chunk) in xf_data.chunks_exact(2).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf);
                }
            }
            // EOF terminates the sheet substream.
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

#[derive(Debug, Default)]
pub(crate) struct SheetHyperlinks {
    pub(crate) hyperlinks: Vec<Hyperlink>,
    pub(crate) warnings: Vec<String>,
}

// Hyperlink record bits (linkOpts / grbit). These come from the MS-XLS HLINK record spec, but we
// treat them as best-effort: Excel files in the wild sometimes contain slightly different flag
// combinations depending on link type.
const HLINK_FLAG_HAS_MONIKER: u32 = 0x0000_0001;
const HLINK_FLAG_HAS_LOCATION: u32 = 0x0000_0008;
const HLINK_FLAG_HAS_DISPLAY: u32 = 0x0000_0010;
const HLINK_FLAG_HAS_TOOLTIP: u32 = 0x0000_0020;
const HLINK_FLAG_HAS_TARGET_FRAME: u32 = 0x0000_0080;

// CLSIDs (COM GUIDs) used by hyperlink monikers.
// GUIDs are stored with the first 3 fields little-endian (standard COM GUID layout).
const CLSID_URL_MONIKER: [u8; 16] = [
    0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
];
const CLSID_FILE_MONIKER: [u8; 16] = [
    0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

// Hardening: cap the number of decoded hyperlinks per worksheet.
//
// XLS files can contain a large number of HLINK records. Without a limit, a crafted file can cause
// large allocations (Vec growth + per-link allocations) and slow imports.
#[cfg(not(test))]
const MAX_HYPERLINKS_PER_SHEET: usize = 50_000;
#[cfg(test)]
const MAX_HYPERLINKS_PER_SHEET: usize = 100;

/// Scan a worksheet BIFF substream for hyperlink records (HLINK, id 0x01B8).
///
/// This is a best-effort parser: malformed records are skipped and surfaced as warnings rather
/// than failing the entire import.
pub(crate) fn parse_biff_sheet_hyperlinks(
    workbook_stream: &[u8],
    start: usize,
    codepage: u16,
) -> Result<SheetHyperlinks, String> {
    let mut out = SheetHyperlinks::default();

    // HLINK records can legally be split across one or more `CONTINUE` records if the hyperlink
    // payload exceeds the BIFF record size limit. Use the logical iterator so we can reassemble
    // those fragments before decoding.
    let allows_continuation = |record_id: u16| record_id == RECORD_HLINK;
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                // Best-effort: stop scanning on malformed record boundaries, but keep any
                // successfully decoded hyperlinks and surface a warning.
                push_warning_bounded(&mut out.warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        // BOF indicates the start of a new substream; stop before consuming the next section so we
        // don't attribute later hyperlinks to this worksheet.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_HLINK => match decode_hlink_record(record.data.as_ref(), codepage) {
                Ok(Some(link)) => {
                    if out.hyperlinks.len() >= MAX_HYPERLINKS_PER_SHEET {
                        out.warnings.push(
                            "too many hyperlinks; additional HLINK records skipped".to_string(),
                        );
                        break;
                    }
                    out.hyperlinks.push(link)
                }
                Ok(None) => {}
                Err(err) => push_warning_bounded(
                    &mut out.warnings,
                    format!(
                        "failed to decode HLINK record at offset {}: {err}",
                        record.offset
                    ),
                ),
            },
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn decode_hlink_record(data: &[u8], codepage: u16) -> Result<Option<Hyperlink>, String> {
    // HLINK [MS-XLS 2.4.110]
    // - ref8 (8 bytes): anchor
    // - guid (16 bytes): hyperlink GUID (ignored)
    // - streamVersion (4 bytes): usually 2
    // - linkOpts (4 bytes): flags
    if data.len() < 32 {
        return Err("HLINK record too short".to_string());
    }

    let rw_first = u16::from_le_bytes([data[0], data[1]]) as u32;
    let rw_last = u16::from_le_bytes([data[2], data[3]]) as u32;
    let col_first = u16::from_le_bytes([data[4], data[5]]) as u32;
    let col_last = u16::from_le_bytes([data[6], data[7]]) as u32;

    if rw_first >= EXCEL_MAX_ROWS
        || rw_last >= EXCEL_MAX_ROWS
        || col_first >= EXCEL_MAX_COLS
        || col_last >= EXCEL_MAX_COLS
    {
        // Ignore out-of-bounds anchors to avoid corrupt coordinates.
        return Ok(None);
    }

    let range = Range::new(
        CellRef::new(rw_first, col_first),
        CellRef::new(rw_last, col_last),
    );

    // Skip guid (16 bytes).
    let mut pos = 8usize + 16usize;

    let stream_version =
        u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    pos += 4;
    if stream_version != 2 {
        // Non-fatal; continue parsing.
        // Some producers may write a different version, but the layout is usually identical.
    }

    let link_opts = u32::from_le_bytes([data[pos], data[pos + 1], data[pos + 2], data[pos + 3]]);
    pos += 4;

    let mut display: Option<String> = None;
    let mut tooltip: Option<String> = None;
    let mut text_mark: Option<String> = None;
    let mut uri: Option<String> = None;

    // Optional: display string.
    if (link_opts & HLINK_FLAG_HAS_DISPLAY) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        display = (!s.is_empty()).then_some(s);
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: target frame (ignored for now).
    if (link_opts & HLINK_FLAG_HAS_TARGET_FRAME) != 0 {
        let (_s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: moniker (external link target).
    if (link_opts & HLINK_FLAG_HAS_MONIKER) != 0 {
        let (parsed_uri, consumed) = parse_hyperlink_moniker(&data[pos..], codepage)?;
        uri = parsed_uri;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: location / text mark (internal target or sub-address).
    if (link_opts & HLINK_FLAG_HAS_LOCATION) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        text_mark = (!s.is_empty()).then_some(s);
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: tooltip.
    if (link_opts & HLINK_FLAG_HAS_TOOLTIP) != 0 {
        let (s, consumed) = parse_hyperlink_string(&data[pos..], codepage)?;
        tooltip = (!s.is_empty()).then_some(s);
        // No further fields depend on the cursor position today, but keep the overflow check so
        // malformed payloads still surface a warning.
        let _ = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    let target = if let Some(mut uri) = uri {
        // Some hyperlink types (notably file links) include both a moniker (base URL/path) and a
        // text mark (sub-address). Preserve that information by encoding the text mark as a URL
        // fragment when possible (e.g. `file:///...#Sheet2!A1`). This matches the common XLSX
        // representation where external hyperlink `Target` may include a fragment.
        if let Some(mark) = text_mark.as_deref() {
            let fragment = mark.trim().trim_start_matches('#');
            if !fragment.is_empty() && !uri.contains('#') {
                uri.push('#');
                uri.push_str(fragment);
            }
        }

        if uri.to_ascii_lowercase().starts_with("mailto:") {
            HyperlinkTarget::Email { uri }
        } else {
            HyperlinkTarget::ExternalUrl { uri }
        }
    } else if let Some(mark) = text_mark.as_deref() {
        let (sheet, cell) = parse_internal_location(mark)
            .ok_or_else(|| "unsupported internal hyperlink".to_string())?;
        HyperlinkTarget::Internal { sheet, cell }
    } else {
        return Err("HLINK record is missing target information".to_string());
    };

    Ok(Some(Hyperlink {
        range,
        target,
        display,
        tooltip,
        rel_id: None,
    }))
}

fn parse_hyperlink_moniker(input: &[u8], codepage: u16) -> Result<(Option<String>, usize), String> {
    if input.len() < 16 {
        return Err("truncated hyperlink moniker".to_string());
    }
    let clsid: [u8; 16] = input[0..16].try_into().expect("slice length verified");

    // URL moniker: UTF-16LE URL with a 32-bit length prefix.
    if clsid == CLSID_URL_MONIKER {
        if input.len() < 20 {
            return Err("truncated URL moniker".to_string());
        }
        let len = u32::from_le_bytes([input[16], input[17], input[18], input[19]]) as usize;
        let mut consumed = 20usize;

        let (url, url_bytes) = parse_utf16_prefixed_string(&input[20..], len)?;
        consumed = consumed
            .checked_add(url_bytes)
            .ok_or_else(|| "URL moniker length overflow".to_string())?;
        return Ok(((!url.is_empty()).then_some(url), consumed));
    }

    // File moniker: best-effort decode to a `file:` URI.
    if clsid == CLSID_FILE_MONIKER {
        // The file moniker payload is more complex (short/long paths, UNC). We attempt a
        // best-effort parse that recovers both the path and the correct payload length so the
        // caller can continue parsing subsequent fields in the HLINK record (tooltip, location,
        // etc) without becoming misaligned.
        //
        // Common layout (best-effort) [MS-OSHARED] / [MS-XLS]:
        //
        //   [cAnti:u32] [ansiPath:cAnti bytes (including NUL)] [endServer:u16] [reserved:u16]
        //   [cbUnicode:u32] [unicodePath:cbUnicode bytes (UTF-16LE, including NUL)]
        if input.len() < 20 {
            return Err("truncated file moniker".to_string());
        }

        let ansi_len = u32::from_le_bytes([input[16], input[17], input[18], input[19]]) as usize;
        if ansi_len > MAX_HLINK_UTF16_BYTES {
            return Err("implausible file moniker ANSI path length".to_string());
        }
        let mut pos = 20usize;

        // ANSI path (8-bit, NUL terminated within the declared byte length).
        let ansi_path = if ansi_len > 0 {
            let end = pos
                .checked_add(ansi_len)
                .ok_or_else(|| "file moniker ANSI length overflow".to_string())?;
            if input.len() < end {
                return Err("truncated file moniker ANSI path".to_string());
            }
            let bytes = &input[pos..end];
            pos = end;

            let nul = bytes.iter().position(|b| *b == 0).unwrap_or(bytes.len());
            let path = strings::decode_ansi(codepage, &bytes[..nul]);
            let path = trim_at_first_nul(path);
            (!path.is_empty()).then_some(path)
        } else {
            None
        };

        // Optional Unicode path extension.
        //
        // Many producers include it; consuming it is important so following HLINK fields parse
        // correctly. Some producers may omit the Unicode tail entirely, so treat it as present only
        // when it looks plausible (otherwise we risk consuming the next HLINK field, e.g. the
        // `location` or `tooltip` string).
        let mut unicode_path: Option<String> = None;
        let Some(unicode_header_end) = pos.checked_add(8) else {
            return Err("file moniker length overflow".to_string());
        };
        if input.len() >= unicode_header_end {
            let end_server = u16::from_le_bytes([input[pos], input[pos + 1]]) as usize;
            // reserved/version (ignored)
            let _reserved = u16::from_le_bytes([input[pos + 2], input[pos + 3]]);
            let unicode_len = u32::from_le_bytes([
                input[pos + 4],
                input[pos + 5],
                input[pos + 6],
                input[pos + 7],
            ]) as usize;

            let available = input.len().saturating_sub(unicode_header_end);
            let end_server_plausible = ansi_len == 0 || end_server <= ansi_len;
            let unicode_len_plausible = unicode_len == 0
                || (unicode_len <= available && unicode_len % 2 == 0)
                || unicode_len
                    .checked_mul(2)
                    .is_some_and(|byte_len| byte_len <= available);

            if end_server_plausible && unicode_len_plausible {
                pos = unicode_header_end;

                if unicode_len > 0 {
                    let (s, consumed) = parse_utf16_prefixed_string(&input[pos..], unicode_len)?;
                    pos = pos
                        .checked_add(consumed)
                        .ok_or_else(|| "file moniker length overflow".to_string())?;

                    let s = trim_at_first_nul(s);
                    if !s.is_empty() {
                        unicode_path = Some(s);
                    }
                }
            }
        }

        let path = unicode_path
            .or(ansi_path)
            .ok_or_else(|| "unsupported file moniker".to_string())?;
        let uri = file_path_to_uri(&path);
        return Ok((Some(uri), pos));
    }

    Err(format!(
        "unsupported hyperlink moniker CLSID {:02X?}",
        clsid
    ))
}

fn percent_encode_uri_path(path: &str) -> String {
    // RFC 3986 path characters: pchar + '/'.
    // pchar = unreserved / pct-encoded / sub-delims / ":" / "@"
    // We percent-encode everything else (including spaces and non-ASCII) for stability.
    fn is_allowed(b: u8) -> bool {
        matches!(
            b,
            b'A'..=b'Z'
                | b'a'..=b'z'
                | b'0'..=b'9'
                | b'-'
                | b'.'
                | b'_'
                | b'~'
                | b'!'
                | b'$'
                | b'&'
                | b'\''
                | b'('
                | b')'
                | b'*'
                | b'+'
                | b','
                | b';'
                | b'='
                | b':'
                | b'@'
                | b'/'
        )
    }

    let mut out = String::with_capacity(path.len());
    for &b in path.as_bytes() {
        if is_allowed(b) {
            out.push(b as char);
        } else {
            out.push('%');
            out.push_str(&format!("{:02X}", b));
        }
    }
    out
}

fn file_path_to_uri(path: &str) -> String {
    let mut p = path.trim().to_string();
    if let Some(idx) = p.find('\0') {
        p.truncate(idx);
    }
    if p.is_empty() {
        return p;
    }

    // If the path already looks like a URI, preserve it.
    if p.contains("://") {
        return p;
    }

    // Handle common Windows extended-length prefixes.
    // - \\?\C:\... => C:\...
    // - \\?\UNC\server\share\... => \\server\share\...
    if let Some(rest) = p.strip_prefix(r"\\?\UNC\") {
        p = format!(r"\\{rest}");
    } else if let Some(rest) = p.strip_prefix(r"\\?\") {
        p = rest.to_string();
    } else if let Some(rest) = p.strip_prefix("//?/UNC/") {
        p = format!("//{rest}");
    } else if let Some(rest) = p.strip_prefix("//?/") {
        p = rest.to_string();
    }

    // Some producers use `C|` instead of `C:` (legacy file URL encoding).
    if p.len() >= 2 {
        let bytes = p.as_bytes();
        if bytes[1] == b'|' && bytes[0].is_ascii_alphabetic() {
            p.replace_range(1..2, ":");
        }
    }

    // Normalize separators to forward slashes and percent-encode unsafe characters.
    let p = p.replace('\\', "/");
    let encoded = percent_encode_uri_path(&p);

    // UNC paths are stored as `\\server\share\...`, which becomes `//server/share/...` after
    // normalization. `file:` + that string yields a valid UNC file URI (`file://server/share/...`).
    if encoded.starts_with("//") {
        return format!("file:{encoded}");
    }

    // Absolute POSIX path.
    if encoded.starts_with('/') {
        return format!("file://{encoded}");
    }

    // Windows drive letter path.
    if encoded.as_bytes().get(1) == Some(&b':') {
        // `file:///C:/path` is represented as `file://` + `/C:/path`.
        return format!("file:///{encoded}");
    }

    // Relative path. Preserve as a relative Target for XLSX compatibility.
    encoded
}

// Keep hyperlink-related string caps consistent across the various encodings we parse.
//
// `parse_hyperlink_string` already caps at 1_000_000 UTF-16 code units. Hyperlink moniker strings
// (URL/file monikers) use a length prefix that producers interpret inconsistently (bytes vs
// characters), so we cap by the maximum number of bytes we might reasonably decode.
const MAX_HLINK_UTF16_CHARS: usize = 1_000_000;
const MAX_HLINK_UTF16_BYTES: usize = MAX_HLINK_UTF16_CHARS * 2;

fn parse_utf16_prefixed_string(input: &[u8], len: usize) -> Result<(String, usize), String> {
    // Heuristic: `len` may be either a byte length or a character count.
    //
    // In BIFF hyperlink structures, producers are inconsistent about whether the length includes:
    // - UTF-16 code units (characters), or
    // - raw bytes.
    //
    // Prefer interpretations that end with a UTF-16 NUL terminator (common in monikers), otherwise
    // fall back to the smaller interpretation to avoid over-consuming into the next field.
    if len == 0 {
        return Ok((String::new(), 0));
    }

    // BIFF hyperlink moniker length fields are untrusted. A corrupt file could declare a huge
    // length and provide enough `CONTINUE` data to satisfy it, causing `decode_utf16le` to try to
    // allocate an enormous `Vec<u16>`. Cap the maximum amount of UTF-16 we will ever decode.
    let len_as_bytes_ok = len <= MAX_HLINK_UTF16_BYTES;
    let len_as_chars_bytes = len.checked_mul(2);
    let len_as_chars_ok =
        len_as_chars_bytes.is_some_and(|byte_len| byte_len <= MAX_HLINK_UTF16_BYTES);

    if !len_as_bytes_ok && !len_as_chars_ok {
        return Err("implausible UTF-16 string length".to_string());
    }

    let mut candidates: Vec<(String, usize, bool)> = Vec::new();

    // Candidate A: `len` as byte length.
    if len_as_bytes_ok && len % 2 == 0 && input.len() >= len {
        let bytes = &input[..len];
        let ends_with_nul = bytes
            .chunks_exact(2)
            .last()
            .is_some_and(|chunk| chunk[0] == 0 && chunk[1] == 0);
        let s = decode_utf16le(bytes)?;
        candidates.push((trim_trailing_nuls(s), len, ends_with_nul));
    }

    // Candidate B: `len` as character count.
    if let Some(byte_len) = len_as_chars_bytes {
        if len_as_chars_ok && byte_len % 2 == 0 && input.len() >= byte_len {
            let bytes = &input[..byte_len];
            let ends_with_nul = bytes
                .chunks_exact(2)
                .last()
                .is_some_and(|chunk| chunk[0] == 0 && chunk[1] == 0);
            let s = decode_utf16le(bytes)?;
            candidates.push((trim_trailing_nuls(s), byte_len, ends_with_nul));
        }
    }

    if candidates.is_empty() {
        return Err("truncated UTF-16 string".to_string());
    }

    // Prefer NUL-terminated candidates; otherwise prefer the shorter byte length.
    candidates.sort_by_key(|(_s, consumed, ends_with_nul)| (!*ends_with_nul, *consumed));
    let (s, consumed, _nul) = candidates.into_iter().next().expect("non-empty candidates");
    Ok((trim_at_first_nul(s), consumed))
}

fn parse_hyperlink_string(input: &[u8], codepage: u16) -> Result<(String, usize), String> {
    // HyperlinkString [MS-XLS 2.5.??]: cch (u32) + UTF-16LE characters.
    // Some producers may store strings as BIFF8 XLUnicodeString; fall back to that on failure.
    if input.len() >= 4 {
        let cch = u32::from_le_bytes([input[0], input[1], input[2], input[3]]) as usize;
        if cch == 0 {
            return Ok((String::new(), 4));
        }
        if cch <= MAX_HLINK_UTF16_CHARS {
            if let Some(byte_len) = cch.checked_mul(2) {
                if input.len() >= 4 + byte_len {
                    let bytes = &input[4..4 + byte_len];
                    let s = decode_utf16le(bytes)?;
                    // Hyperlink-related strings in BIFF are frequently NUL terminated, but we've
                    // observed files in the wild that include embedded NULs + trailing garbage
                    // within the declared length. Truncate at the first NUL for best-effort
                    // compatibility (mirrors how file moniker paths are handled).
                    let s = trim_at_first_nul(trim_trailing_nuls(s));
                    return Ok((s, 4 + byte_len));
                }
            }
        }
    }

    // Fallback: BIFF8 XLUnicodeString (u16 length + flags).
    let (s, consumed) = strings::parse_biff8_unicode_string(input, codepage)?;
    Ok((trim_at_first_nul(trim_trailing_nuls(s)), consumed))
}

fn decode_utf16le(bytes: &[u8]) -> Result<String, String> {
    if bytes.len() % 2 != 0 {
        return Err("truncated UTF-16 string".to_string());
    }
    let mut u16s = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    Ok(String::from_utf16_lossy(&u16s))
}

fn trim_trailing_nuls(mut s: String) -> String {
    while s.ends_with('\0') {
        s.pop();
    }
    s
}

fn trim_at_first_nul(mut s: String) -> String {
    if let Some(idx) = s.find('\0') {
        s.truncate(idx);
    }
    s
}

fn parse_internal_location(location: &str) -> Option<(String, CellRef)> {
    // Mirrors the XLSX hyperlink `location` parsing logic.
    let mut loc = location.trim();
    if let Some(rest) = loc.strip_prefix('#') {
        loc = rest;
    }

    let (sheet, cell) = loc.split_once('!')?;
    let sheet = unquote_sheet_name(sheet.trim());

    let cell_str = cell.trim();
    let cell_str = cell_str
        .split_once(':')
        .map(|(start, _)| start)
        .unwrap_or(cell_str);
    let cell = CellRef::from_a1(cell_str).ok()?;
    Some((sheet, cell))
}

fn unquote_sheet_name(name: &str) -> String {
    // Excel quotes sheet names with single quotes; embedded quotes are doubled.
    let mut s = name.trim();
    if s.starts_with('\'') && s.ends_with('\'') && s.len() >= 2 {
        s = &s[1..s.len() - 1];
        return s.replace("''", "'");
    }
    s.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeSet;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    #[test]
    fn utf16_prefixed_string_rejects_implausible_length_small_input() {
        let input = [0u8; 8];
        let err = parse_utf16_prefixed_string(&input, u32::MAX as usize).unwrap_err();
        assert!(
            err.contains("implausible"),
            "expected implausible-length error, got {err:?}"
        );
    }

    #[test]
    fn utf16_prefixed_string_rejects_implausible_length_even_when_buffer_is_large_enough() {
        let len = MAX_HLINK_UTF16_BYTES + 2;
        let input = vec![0u8; len];
        let err = parse_utf16_prefixed_string(&input, len).unwrap_err();
        assert!(
            err.contains("implausible"),
            "expected implausible-length error, got {err:?}"
        );
    }

    #[test]
    fn utf16_prefixed_string_reports_truncated_when_only_bytes_interpretation_is_plausible() {
        // `len` fits within the byte cap but exceeds the char->bytes cap, so only the byte-length
        // interpretation is considered. Since the buffer is too small, we should report truncation
        // rather than "implausible length".
        let len = MAX_HLINK_UTF16_CHARS + 2;
        let err = parse_utf16_prefixed_string(&[], len).unwrap_err();
        assert!(
            err.contains("truncated"),
            "expected truncated error, got {err:?}"
        );
    }

    #[test]
    fn file_moniker_rejects_implausible_ansi_path_length() {
        let mut input = Vec::new();
        input.extend_from_slice(&CLSID_FILE_MONIKER);
        input.extend_from_slice(&((MAX_HLINK_UTF16_BYTES + 1) as u32).to_le_bytes());
        let err = parse_hyperlink_moniker(&input, 1252).unwrap_err();
        assert!(
            err.contains("implausible file moniker ANSI path length"),
            "expected implausible ANSI length error, got {err:?}"
        );
    }

    #[test]
    fn sheet_row_col_scan_stops_on_truncated_record() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // ROW 1 with explicit height = 20.0 points (400 twips).
        let mut row_payload = [0u8; 16];
        row_payload[0..2].copy_from_slice(&1u16.to_le_bytes());
        row_payload[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_record = record(RECORD_ROW, &row_payload);

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes

        let stream = [sheet_bof, row_record, truncated].concat();
        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert_eq!(props.rows.get(&1).and_then(|p| p.height), Some(20.0));
    }

    #[test]
    fn sheet_row_col_warnings_are_bounded() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Emit many malformed ROW records (payload < 16 bytes). Without bounding this would
        // allocate unbounded warning strings.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 100) {
            stream.extend_from_slice(&record(RECORD_ROW, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert_eq!(props.warnings.len(), MAX_WARNINGS_PER_SHEET + 1);
        assert_eq!(
            props
                .warnings
                .iter()
                .filter(|w| w.contains(WARNINGS_SUPPRESSED_MESSAGE))
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            props.warnings
        );
    }

    #[test]
    fn sheet_row_col_sort_warnings_are_bounded() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Emit many malformed SORT records (payload < 24 bytes).
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 100) {
            stream.extend_from_slice(&record(RECORD_SORT, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert_eq!(props.warnings.len(), MAX_WARNINGS_PER_SHEET + 1);
        assert_eq!(
            props
                .warnings
                .iter()
                .filter(|w| w.contains(WARNINGS_SUPPRESSED_MESSAGE))
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            props.warnings
        );
    }

    #[test]
    fn clamps_colinfo_column_ranges_to_excel_bounds() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // COLINFO with a range that extends past the model's max column.
        let first_col = EXCEL_MAX_COLS - 2;
        let last_col = EXCEL_MAX_COLS + 2;

        let mut colinfo_payload = Vec::new();
        colinfo_payload.extend_from_slice(&(first_col as u16).to_le_bytes());
        colinfo_payload.extend_from_slice(&(last_col as u16).to_le_bytes());
        colinfo_payload.extend_from_slice(&256u16.to_le_bytes()); // width = 1.0
        colinfo_payload.extend_from_slice(&0u16.to_le_bytes()); // ixfe
        colinfo_payload.extend_from_slice(&0u16.to_le_bytes()); // options
        colinfo_payload.extend_from_slice(&0u16.to_le_bytes()); // reserved

        let stream = [
            sheet_bof,
            record(RECORD_COLINFO, &colinfo_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");

        assert_eq!(
            props.cols.keys().copied().collect::<Vec<_>>(),
            vec![EXCEL_MAX_COLS - 2, EXCEL_MAX_COLS - 1]
        );
        assert!(!props.cols.contains_key(&EXCEL_MAX_COLS));
        assert!(!props.cols.contains_key(&(EXCEL_MAX_COLS + 1)));
        assert!(!props.cols.contains_key(&(EXCEL_MAX_COLS + 2)));

        assert!(
            props
                .warnings
                .iter()
                .any(|w| w.contains("COLINFO") && w.contains("truncated")),
            "expected truncation warning, got {:?}",
            props.warnings
        );
    }

    #[test]
    fn sheet_hyperlink_warnings_are_bounded() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Emit many malformed HLINK records (payload < 32 bytes). Each record should surface a
        // warning, but the vector is capped.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 100) {
            stream.extend_from_slice(&record(RECORD_HLINK, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let links = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");
        assert!(links.hyperlinks.is_empty());
        assert_eq!(links.warnings.len(), MAX_WARNINGS_PER_SHEET + 1);
        assert_eq!(
            links
                .warnings
                .iter()
                .filter(|w| w.contains(WARNINGS_SUPPRESSED_MESSAGE))
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            links.warnings
        );
    }

    #[test]
    fn parses_sheet_cell_xf_indices_including_mul_records() {
        // NUMBER cell (A1) with xf=3.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes()); // row
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes()); // col
        number_payload[4..6].copy_from_slice(&3u16.to_le_bytes()); // xf

        // MULBLANK row=1, cols 0..2 with xf {10,11,12}.
        let mut mulblank_payload = Vec::new();
        mulblank_payload.extend_from_slice(&1u16.to_le_bytes()); // row
        mulblank_payload.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        mulblank_payload.extend_from_slice(&10u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&11u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&12u16.to_le_bytes());
        mulblank_payload.extend_from_slice(&2u16.to_le_bytes()); // colLast

        // MULRK row=2, cols 1..2 with xf {20,21}.
        let mut mulrk_payload = Vec::new();
        mulrk_payload.extend_from_slice(&2u16.to_le_bytes()); // row
        mulrk_payload.extend_from_slice(&1u16.to_le_bytes()); // colFirst
                                                              // cell 1: xf=20 + dummy rk value
        mulrk_payload.extend_from_slice(&20u16.to_le_bytes());
        mulrk_payload.extend_from_slice(&0u32.to_le_bytes());
        // cell 2: xf=21 + dummy rk value
        mulrk_payload.extend_from_slice(&21u16.to_le_bytes());
        mulrk_payload.extend_from_slice(&0u32.to_le_bytes());
        mulrk_payload.extend_from_slice(&2u16.to_le_bytes()); // colLast

        let stream = [
            record(RECORD_NUMBER, &number_payload),
            record(RECORD_MULBLANK, &mulblank_payload),
            record(RECORD_MULRK, &mulrk_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(3));
        assert_eq!(xfs.get(&CellRef::new(1, 0)).copied(), Some(10));
        assert_eq!(xfs.get(&CellRef::new(1, 1)).copied(), Some(11));
        assert_eq!(xfs.get(&CellRef::new(1, 2)).copied(), Some(12));
        assert_eq!(xfs.get(&CellRef::new(2, 1)).copied(), Some(20));
        assert_eq!(xfs.get(&CellRef::new(2, 2)).copied(), Some(21));
    }

    #[test]
    fn parses_mergedcells_records() {
        // First record: A1:B1.
        let mut merged1 = Vec::new();
        merged1.extend_from_slice(&1u16.to_le_bytes()); // cAreas
        merged1.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        merged1.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        merged1.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        merged1.extend_from_slice(&1u16.to_le_bytes()); // colLast

        // Second record: one valid area (C2:D3) and one out-of-bounds (colFirst >= EXCEL_MAX_COLS).
        let mut merged2 = Vec::new();
        merged2.extend_from_slice(&2u16.to_le_bytes()); // cAreas
                                                        // C2:D3 => rows 1..2, cols 2..3 (0-based)
        merged2.extend_from_slice(&1u16.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&2u16.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&2u16.to_le_bytes()); // colFirst
        merged2.extend_from_slice(&3u16.to_le_bytes()); // colLast
                                                        // Out-of-bounds cols.
        merged2.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&0u16.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colFirst (OOB)
        merged2.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colLast (OOB)

        let stream = [
            record(RECORD_MERGEDCELLS, &merged1),
            record(RECORD_MERGEDCELLS, &merged2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let ranges = parse_biff_sheet_merged_cells(&stream, 0).expect("parse");
        assert_eq!(
            ranges,
            vec![
                Range::from_a1("A1:B1").unwrap(),
                Range::from_a1("C2:D3").unwrap(),
            ]
        );
    }

    #[test]
    fn parses_number_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&1u16.to_le_bytes()); // row
        data.extend_from_slice(&2u16.to_le_bytes()); // col
        data.extend_from_slice(&7u16.to_le_bytes()); // xf
        data.extend_from_slice(&0f64.to_le_bytes()); // value

        let stream = [
            record(RECORD_NUMBER, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(1, 2)).copied(), Some(7));
    }

    #[test]
    fn parses_rk_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&3u16.to_le_bytes()); // row
        data.extend_from_slice(&4u16.to_le_bytes()); // col
        data.extend_from_slice(&9u16.to_le_bytes()); // xf
        data.extend_from_slice(&0u32.to_le_bytes()); // rk

        let stream = [record(RECORD_RK, &data), record(records::RECORD_EOF, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(3, 4)).copied(), Some(9));
    }

    #[test]
    fn parses_blank_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&10u16.to_le_bytes()); // row
        data.extend_from_slice(&3u16.to_le_bytes()); // col
        data.extend_from_slice(&2u16.to_le_bytes()); // xf

        let stream = [
            record(RECORD_BLANK, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(10, 3)).copied(), Some(2));
    }

    #[test]
    fn parses_labelsst_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&0u16.to_le_bytes()); // col
        data.extend_from_slice(&55u16.to_le_bytes()); // xf
        data.extend_from_slice(&123u32.to_le_bytes()); // sst index

        let stream = [
            record(RECORD_LABELSST, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(55));
    }

    #[test]
    fn parses_label_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&2u16.to_le_bytes()); // row
        data.extend_from_slice(&1u16.to_le_bytes()); // col
        data.extend_from_slice(&77u16.to_le_bytes()); // xf
        data.extend_from_slice(&0u16.to_le_bytes()); // cch (placeholder)

        let stream = [
            record(RECORD_LABEL_BIFF5, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(2, 1)).copied(), Some(77));
    }

    #[test]
    fn parses_boolerr_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&9u16.to_le_bytes()); // row
        data.extend_from_slice(&8u16.to_le_bytes()); // col
        data.extend_from_slice(&5u16.to_le_bytes()); // xf
        data.push(1); // value
        data.push(0); // fErr

        let stream = [
            record(RECORD_BOOLERR, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(9, 8)).copied(), Some(5));
    }

    #[test]
    fn parses_formula_record_ixfe() {
        let mut data = Vec::new();
        data.extend_from_slice(&4u16.to_le_bytes()); // row
        data.extend_from_slice(&4u16.to_le_bytes()); // col
        data.extend_from_slice(&6u16.to_le_bytes()); // xf
        data.extend_from_slice(&[0u8; 14]); // rest of FORMULA record (dummy)

        let stream = [
            record(RECORD_FORMULA, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(4, 4)).copied(), Some(6));
    }

    #[test]
    fn prefers_last_record_for_duplicate_cells() {
        let blank = {
            let mut data = Vec::new();
            data.extend_from_slice(&0u16.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&1u16.to_le_bytes()); // xf
            record(RECORD_BLANK, &data)
        };

        let number = {
            let mut data = Vec::new();
            data.extend_from_slice(&0u16.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&2u16.to_le_bytes()); // xf
            data.extend_from_slice(&0f64.to_le_bytes());
            record(RECORD_NUMBER, &data)
        };

        let stream = [blank, number, record(records::RECORD_EOF, &[])].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(2));
    }

    #[test]
    fn skips_out_of_bounds_cells() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes()); // row
        data.extend_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // col (out of bounds)
        data.extend_from_slice(&1u16.to_le_bytes()); // xf

        let stream = [
            record(RECORD_BLANK, &data),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert!(xfs.is_empty());
    }

    #[test]
    fn sheet_row_col_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // ROW 1 with explicit height = 20.0 points (400 twips).
        let mut row_payload = [0u8; 16];
        row_payload[0..2].copy_from_slice(&1u16.to_le_bytes());
        row_payload[6..8].copy_from_slice(&400u16.to_le_bytes());
        let row_record = record(RECORD_ROW, &row_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        let stream = [sheet_bof, row_record, next_bof].concat();
        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert_eq!(props.rows.get(&1).and_then(|p| p.height), Some(20.0));
    }

    #[test]
    fn sheet_cell_xf_scan_stops_at_next_bof_without_eof() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // NUMBER cell at (0,0) with xf=7.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        number_payload[4..6].copy_from_slice(&7u16.to_le_bytes());
        let number_record = record(RECORD_NUMBER, &number_payload);

        // BOF for the next substream; no EOF record for the worksheet.
        let next_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        let stream = [sheet_bof, number_record, next_bof].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(7));
    }

    #[test]
    fn sheet_cell_xf_scan_stops_on_truncated_record() {
        let sheet_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);

        // NUMBER cell at (0,0) with xf=7.
        let mut number_payload = vec![0u8; 14];
        number_payload[0..2].copy_from_slice(&0u16.to_le_bytes());
        number_payload[2..4].copy_from_slice(&0u16.to_le_bytes());
        number_payload[4..6].copy_from_slice(&7u16.to_le_bytes());
        let number_record = record(RECORD_NUMBER, &number_payload);

        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes

        let stream = [sheet_bof, number_record, truncated].concat();
        let xfs = parse_biff_sheet_cell_xf_indices_filtered(&stream, 0, None).expect("parse");
        assert_eq!(xfs.get(&CellRef::new(0, 0)).copied(), Some(7));
    }

    #[test]
    fn parses_sheet_protection_records() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_PASSWORD, &0xCBEBu16.to_le_bytes()),
            record(RECORD_OBJPROTECT, &0u16.to_le_bytes()),
            record(RECORD_SCENPROTECT, &0u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(parsed.protection.enabled, true);
        assert_eq!(parsed.protection.password_hash, Some(0xCBEB));
        assert_eq!(parsed.protection.edit_objects, true);
        assert_eq!(parsed.protection.edit_scenarios, true);
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_password_hash_zero_is_none() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            // Hash value 0 indicates "no password" in Excel's legacy protection scheme.
            record(RECORD_PASSWORD, &0u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(parsed.protection.enabled, true);
        assert_eq!(parsed.protection.password_hash, None);
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_scan_stops_at_next_bof() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            // BOF for the next substream; no EOF record for the worksheet protection scan.
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            // This record should be ignored because it's in the next substream.
            record(RECORD_PROTECT, &0u16.to_le_bytes()),
        ]
        .concat();
        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(parsed.protection.enabled, true);
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_warns_on_truncated_records_and_continues() {
        // Emit truncated protection records followed by valid ones; parser should warn but still
        // return the final values.
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_PROTECT, &[1]), // truncated
            record(RECORD_PROTECT, &1u16.to_le_bytes()),
            record(RECORD_PASSWORD, &[0xEF]), // truncated
            record(RECORD_PASSWORD, &0xBEEFu16.to_le_bytes()),
            record(RECORD_OBJPROTECT, &[0]), // truncated
            record(RECORD_OBJPROTECT, &0u16.to_le_bytes()),
            record(RECORD_SCENPROTECT, &[0]), // truncated
            record(RECORD_SCENPROTECT, &0u16.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();
        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(parsed.protection.enabled, true);
        assert_eq!(parsed.protection.password_hash, Some(0xBEEF));
        assert_eq!(parsed.protection.edit_objects, true);
        assert_eq!(parsed.protection.edit_scenarios, true);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated PROTECT record")),
            "expected truncated-PROTECT warning, got {:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated PASSWORD record")),
            "expected truncated-PASSWORD warning, got {:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated OBJPROTECT record")),
            "expected truncated-OBJPROTECT warning, got {:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated SCENPROTECT record")),
            "expected truncated-SCENPROTECT warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_manual_page_breaks_and_warns_on_truncated_entries() {
        // HorizontalPageBreaks with cbrk=2 but only one complete HorzBrk entry (6 bytes).
        let mut horizontal = Vec::new();
        horizontal.extend_from_slice(&2u16.to_le_bytes()); // cbrk
        horizontal.extend_from_slice(&2u16.to_le_bytes()); // row (first row below break) => break after 1
        horizontal.extend_from_slice(&0u16.to_le_bytes()); // colStart
        horizontal.extend_from_slice(&255u16.to_le_bytes()); // colEnd

        // VerticalPageBreaks with a single complete entry.
        let mut vertical = Vec::new();
        vertical.extend_from_slice(&1u16.to_le_bytes()); // cbrk
        vertical.extend_from_slice(&3u16.to_le_bytes()); // col (first col right of break) => break after 2
        vertical.extend_from_slice(&0u16.to_le_bytes()); // rowStart
        vertical.extend_from_slice(&0u16.to_le_bytes()); // rowEnd

        let stream = [
            record(RECORD_HORIZONTALPAGEBREAKS, &horizontal),
            record(RECORD_VERTICALPAGEBREAKS, &vertical),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_manual_page_breaks(&stream, 0).expect("parse");
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("HorizontalPageBreaks")),
            "expected HorizontalPageBreaks warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.manual_page_breaks.row_breaks_after,
            BTreeSet::from([1u32])
        );
        assert_eq!(
            parsed.manual_page_breaks.col_breaks_after,
            BTreeSet::from([2u32])
        );
    }

    fn setup_payload(
        i_paper_size: u16,
        i_scale: u16,
        i_fit_width: u16,
        i_fit_height: u16,
        grbit: u16,
        num_hdr: f64,
        num_ftr: f64,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&i_paper_size.to_le_bytes());
        out.extend_from_slice(&i_scale.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // iPageStart
        out.extend_from_slice(&i_fit_width.to_le_bytes());
        out.extend_from_slice(&i_fit_height.to_le_bytes());
        out.extend_from_slice(&grbit.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // iRes
        out.extend_from_slice(&0u16.to_le_bytes()); // iVRes
        out.extend_from_slice(&num_hdr.to_le_bytes());
        out.extend_from_slice(&num_ftr.to_le_bytes());
        out.extend_from_slice(&1u16.to_le_bytes()); // iCopies
        out
    }

    #[test]
    fn parses_page_setup_margins_and_fit_to_page_scaling() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(
                    9,      // A4
                    77,     // iScale (ignored when fit-to-page)
                    2,      // iFitWidth
                    3,      // iFitHeight
                    0x0000, // landscape (fPortrait=0)
                    0.5,    // header inches
                    0.6,    // footer inches
                ),
            ),
            record(RECORD_LEFTMARGIN, &1.0f64.to_le_bytes()),
            record(RECORD_LEFTMARGIN, &2.0f64.to_le_bytes()), // last wins
            record(RECORD_RIGHTMARGIN, &1.2f64.to_le_bytes()),
            record(RECORD_TOPMARGIN, &1.3f64.to_le_bytes()),
            record(RECORD_BOTTOMMARGIN, &1.4f64.to_le_bytes()),
            record(RECORD_WSBOOL, &0x0100u16.to_le_bytes()), // fFitToPage=1
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed.page_setup.expect("expected page_setup");
        assert_eq!(setup.paper_size.code, 9);
        assert_eq!(setup.orientation, Orientation::Landscape);
        assert_eq!(setup.scaling, Scaling::FitTo { width: 2, height: 3 });
        assert_eq!(setup.margins.left, 2.0);
        assert_eq!(setup.margins.right, 1.2);
        assert_eq!(setup.margins.top, 1.3);
        assert_eq!(setup.margins.bottom, 1.4);
        assert_eq!(setup.margins.header, 0.5);
        assert_eq!(setup.margins.footer, 0.6);
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_percent_scaling_when_fit_to_page_disabled() {
        let grbit = 0x0002u16; // fPortrait=1
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(1, 80, 1, 1, grbit, 0.3, 0.3), // iScale=80%
            ),
            record(RECORD_WSBOOL, &0u16.to_le_bytes()), // fFitToPage=0
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed.page_setup.expect("expected page_setup");
        assert_eq!(setup.scaling, Scaling::Percent(80));
    }

    #[test]
    fn setup_f_nopls_ignores_printer_fields() {
        // fNoPls=1 => iPaperSize/iScale/fPortrait are undefined and must be ignored.
        let grbit = 0x0004u16; // fNoPls
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(
                RECORD_SETUP,
                &setup_payload(9, 80, 1, 1, grbit, 0.4, 0.5), // values ignored except header/footer
            ),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed.page_setup.expect("expected page_setup");
        assert_eq!(setup.paper_size, PageSetup::default().paper_size);
        assert_eq!(setup.orientation, Orientation::Portrait);
        assert_eq!(setup.scaling, Scaling::Percent(100));
        assert_eq!(setup.margins.header, 0.4);
        assert_eq!(setup.margins.footer, 0.5);
    }

    #[test]
    fn warns_on_truncated_margin_records_and_continues() {
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_TOPMARGIN, &[0xAA, 0xBB]), // truncated
            record(RECORD_LEFTMARGIN, &1.0f64.to_le_bytes()),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_print_settings(&stream, 0).expect("parse");
        let setup = parsed.page_setup.expect("expected page_setup");
        assert_eq!(setup.margins.left, 1.0);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated TOPMARGIN record")),
            "expected truncated-TOPMARGIN warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_hyperlink_scan_caps_output_and_stops_early() {
        fn internal_hlink_payload(location: &str) -> Vec<u8> {
            let mut data = Vec::new();

            // ref8 anchor: A1 (0-based row/col).
            data.extend_from_slice(&0u16.to_le_bytes()); // rwFirst
            data.extend_from_slice(&0u16.to_le_bytes()); // rwLast
            data.extend_from_slice(&0u16.to_le_bytes()); // colFirst
            data.extend_from_slice(&0u16.to_le_bytes()); // colLast

            // guid (ignored).
            data.extend_from_slice(&[0u8; 16]);

            // streamVersion + linkOpts.
            data.extend_from_slice(&2u32.to_le_bytes());
            data.extend_from_slice(&HLINK_FLAG_HAS_LOCATION.to_le_bytes());

            // HyperlinkString (u32 char count + UTF-16LE bytes).
            let u16s: Vec<u16> = location.encode_utf16().collect();
            data.extend_from_slice(&(u16s.len() as u32).to_le_bytes());
            for ch in u16s {
                data.extend_from_slice(&ch.to_le_bytes());
            }

            data
        }

        let mut parts: Vec<Vec<u8>> = Vec::new();
        parts.push(record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Emit more HLINK records than the hard cap; the parser should stop at the cap and emit a
        // truncation warning.
        for _ in 0..(MAX_HYPERLINKS_PER_SHEET + 1) {
            let payload = internal_hlink_payload("Sheet1!A1");
            parts.push(record(RECORD_HLINK, &payload));
        }

        // Malformed record after the extra hyperlink; should be ignored because the scanner stops
        // early once the cap is reached.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x0001u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[1, 2]); // missing 2 bytes
        parts.push(truncated);

        let stream = parts.concat();
        let parsed = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");

        assert_eq!(parsed.hyperlinks.len(), MAX_HYPERLINKS_PER_SHEET);
        assert_eq!(
            parsed.warnings,
            vec!["too many hyperlinks; additional HLINK records skipped".to_string()]
        );
    }
}
