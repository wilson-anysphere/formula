use std::collections::{BTreeMap, HashMap};

use formula_model::{
    autofilter::{
        FilterColumn, FilterCriterion, FilterJoin, FilterValue, SortCondition, SortState,
    },
    CellRef, Hyperlink, HyperlinkTarget, ManualPageBreaks, OutlinePr, Range, SheetPane,
    SheetProtection, SheetSelection, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};

use super::records;
use super::rgce;
use super::strings;
use super::worksheet_formulas;

/// Hard cap on the number of per-cell XF indices tracked per worksheet.
///
/// Worksheet cell-record scans may encounter large numbers of formatted-but-empty
/// cells (notably `BLANK`/`MULBLANK`). Without a cap a crafted workbook can force
/// unbounded `HashMap<CellRef, u16>` allocations, risking OOM.
const MAX_CELL_XF_ENTRIES_PER_SHEET: usize = 1_000_000;

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
/// ContinueFrt12 [MS-XLS] (Future Record Type continuation; BIFF8 only)
///
/// Large FRT records (including AutoFilter12/Sort12/SortData12) may be continued across one or more
/// `ContinueFrt12` records. These records also begin with an `FrtHeader`; the bytes after the
/// header should be appended to the previous FRT record payload.
const RECORD_CONTINUEFRT12: u16 = 0x087F;
/// Sort12 (Future Record Type; BIFF8 only)
const RECORD_SORT12: u16 = 0x0890;
// Alternate Sort12 rt value observed in some non-Excel producers.
const RECORD_SORT12_ALT: u16 = 0x0880;
/// SortData12 (Future Record Type; BIFF8 only)
const RECORD_SORTDATA12: u16 = 0x0895;
// Alternate SortData12 rt value observed in some non-Excel producers.
const RECORD_SORTDATA12_ALT: u16 = 0x0881;
const RECORD_WSBOOL: u16 = 0x0081;
/// MERGEDCELLS [MS-XLS 2.4.139]
const RECORD_MERGEDCELLS: u16 = 0x00E5;

/// Maximum merged ranges to parse from a single sheet BIFF substream.
///
/// Malformed `.xls` files can contain a huge number of `MERGEDCELLS` records (or very large
/// `cAreas` counts), which would otherwise cause unbounded growth of the merged-range vector.
#[cfg(not(test))]
const MAX_MERGED_RANGES_PER_SHEET: usize = 100_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_MERGED_RANGES_PER_SHEET: usize = 1_000;

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
/// TABLE [MS-XLS 2.4.313]
///
/// Used for What-If Analysis data tables (the legacy `TABLE()` function), referenced by `PtgTbl`
/// tokens inside FORMULA `rgce` streams.
const RECORD_TABLE: u16 = 0x0236;

/// Scan a worksheet BIFF substream for string cell records (`LABELSST`, id `0x00FD`) and return a
/// mapping from cell address to SST index.
///
/// This is used to associate workbook-global SST metadata (e.g. phonetic guides stored in
/// `XLUnicodeRichExtendedString.ExtRst`) with individual worksheet cells.
///
/// `LABELSST` payload layout ([MS-XLS] 2.4.148):
/// - `rw: u16`
/// - `col: u16`
/// - `ixfe: u16` (ignored)
/// - `isst: u32` (SST index)
///
/// Best-effort: malformed/truncated records are skipped.
pub(crate) fn parse_biff_sheet_labelsst_indices(
    workbook_stream: &[u8],
    start: usize,
    sst_phonetics: Option<&[Option<String>]>,
) -> Result<SheetLabelSstIndices, String> {
    let mut out = SheetLabelSstIndices::default();

    let mut scanned = 0usize;
    for record in records::BestEffortSubstreamIter::from_offset(workbook_stream, start)? {
        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_warning_bounded_force(
                    &mut out.warnings,
                    "record counter overflow while scanning LABELSST indices; stopping early",
                );
                break;
            }
        };
        if scanned > MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN {
            push_warning_bounded_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning LABELSST indices (cap={MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN}); stopping early"
                ),
            );
            break;
        }

        match record.record_id {
            RECORD_LABELSST => {
                let data = record.data;
                if data.len() < 10 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "malformed LABELSST record at offset {}: expected >=10 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }

                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let isst = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);

                if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
                    continue;
                }

                if let Some(phonetics) = sst_phonetics {
                    let idx = isst as usize;
                    if phonetics.get(idx).and_then(|p| p.as_ref()).is_none() {
                        continue;
                    }
                }

                let cell = CellRef::new(row, col);
                if out.indices.len() >= MAX_LABELSST_ENTRIES_PER_SHEET
                    && !out.indices.contains_key(&cell)
                {
                    push_warning_bounded_force(
                        &mut out.warnings,
                        format!(
                            "too many LABELSST indices (cap={MAX_LABELSST_ENTRIES_PER_SHEET}); stopping early"
                        ),
                    );
                    break;
                }

                out.indices.insert(cell, isst);
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

#[derive(Debug, Default)]
pub(crate) struct SheetLabelSstIndices {
    pub(crate) indices: HashMap<CellRef, u32>,
    pub(crate) warnings: Vec<String>,
}

/// Hard cap on the number of `LABELSST` cell records retained per sheet.
///
/// This bounds memory usage for `.xls` files with extremely large numbers of string cells.
#[cfg(not(test))]
const MAX_LABELSST_ENTRIES_PER_SHEET: usize = 1_000_000;
#[cfg(test)]
const MAX_LABELSST_ENTRIES_PER_SHEET: usize = 256;

/// Hard cap on the number of BIFF records scanned while searching for `LABELSST` records.
///
/// The `.xls` importer may run multiple best-effort passes over a sheet stream. Without a cap, a
/// crafted workbook can force excessive work by making these scans traverse huge substreams.
#[cfg(not(test))]
const MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN: usize = 500_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN: usize = 1_000;

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

/// Hard cap on the number of selection ranges decoded from a single `SELECTION` record.
///
/// The `SELECTION` record can declare up to 65k ranges (`cref: u16`). Without a cap a crafted file
/// can force large `Vec<Range>` allocations during the best-effort worksheet view-state scan.
#[cfg(not(test))]
const MAX_SELECTION_RANGES_PER_RECORD: usize = 4_096;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_SELECTION_RANGES_PER_RECORD: usize = 64;

/// Hard cap on the number of distinct `SELECTION` records retained during view-state scanning.
///
/// Excel typically emits at most one selection per pane, but malformed files can include an
/// unbounded number of `SELECTION` records. We only need a small set to pick the active selection.
#[cfg(not(test))]
const MAX_SELECTION_RECORDS_PER_SHEET_VIEW_STATE: usize = 16;
#[cfg(test)]
const MAX_SELECTION_RECORDS_PER_SHEET_VIEW_STATE: usize = 8;

// Manual page breaks (worksheet substream).
// - VERTICALPAGEBREAKS: [MS-XLS 2.4.349]
// - HORIZONTALPAGEBREAKS: [MS-XLS 2.4.115]
const RECORD_VERTICALPAGEBREAKS: u16 = 0x001A;
const RECORD_HORIZONTALPAGEBREAKS: u16 = 0x001B;

/// Hard cap on the number of BIFF records scanned while searching for manual page breaks.
///
/// The `.xls` importer performs multiple best-effort passes over each worksheet substream. Without
/// a cap, a crafted workbook with millions of cell records can force excessive work even when a
/// particular feature (like page breaks) is absent.
#[cfg(not(test))]
const MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN: usize = 500_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN: usize = 1_000;

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

/// Cap warnings collected by worksheet *metadata* scans (view-state/protection).
///
/// These scans are best-effort and intentionally resilient to malformed records. Without a cap, a
/// crafted `.xls` can allocate an unbounded number of warning strings.
const MAX_WARNINGS_PER_SHEET_METADATA: usize = MAX_WARNINGS_PER_SHEET;
const SHEET_METADATA_WARNINGS_SUPPRESSED: &str = "additional sheet metadata warnings suppressed";

/// Hard cap on the number of BIFF records scanned during sheet metadata passes.
///
/// View-state/protection scans are best-effort and should not traverse arbitrarily large worksheets
/// (which may contain millions of cell records). Without a cap, a crafted file can force excessive
/// work by making the importer scan huge substreams multiple times.
#[cfg(not(test))]
const MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN: usize = 500_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN: usize = 1_000;

fn push_sheet_metadata_warning(warnings: &mut Vec<String>, warning: impl Into<String>) {
    if warnings.len() < MAX_WARNINGS_PER_SHEET_METADATA {
        warnings.push(warning.into());
        return;
    }
    // Add a single terminal warning so callers have a hint that the import was noisy.
    if warnings.len() == MAX_WARNINGS_PER_SHEET_METADATA {
        warnings.push(SHEET_METADATA_WARNINGS_SUPPRESSED.to_string());
    }
}

/// Push a sheet-metadata warning but ensure it is present even if the warning buffer is already full.
///
/// This mirrors [`push_warning_bounded_force`] for the view-state/protection warning buffer so we
/// can surface "critical" metadata warnings (like hardening caps) even when earlier parsing already
/// exhausted the warning budget.
fn push_sheet_metadata_warning_force(warnings: &mut Vec<String>, warning: impl Into<String>) {
    let warning = warning.into();

    if warnings.len() < MAX_WARNINGS_PER_SHEET_METADATA {
        warnings.push(warning);
        return;
    }

    let replace_idx = if warnings.len() == MAX_WARNINGS_PER_SHEET_METADATA + 1
        && warnings
            .last()
            .is_some_and(|w| w == SHEET_METADATA_WARNINGS_SUPPRESSED)
    {
        MAX_WARNINGS_PER_SHEET_METADATA - 1
    } else {
        warnings.len() - 1
    };

    if let Some(slot) = warnings.get_mut(replace_idx) {
        *slot = warning;
    } else {
        // Should be unreachable, but fall back to the bounded helper for safety.
        push_sheet_metadata_warning(warnings, warning);
    }
}

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

/// Push a warning but ensure it is present even if the warning buffer is already full.
///
/// This is used for "critical" warnings (e.g. hardening caps) where we want to surface the
/// condition even if earlier best-effort parsing already exhausted the warning budget.
fn push_warning_bounded_force(warnings: &mut Vec<String>, warning: impl Into<String>) {
    let warning = warning.into();

    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(warning);
        return;
    }

    // Keep the warning buffer size bounded. Prefer preserving the terminal
    // `WARNINGS_SUPPRESSED_MESSAGE` marker (when present) and replace the last "real" warning.
    let replace_idx = if warnings.len() == MAX_WARNINGS_PER_SHEET + 1
        && warnings
            .last()
            .is_some_and(|w| w == WARNINGS_SUPPRESSED_MESSAGE)
    {
        MAX_WARNINGS_PER_SHEET - 1
    } else {
        warnings.len() - 1
    };

    if let Some(slot) = warnings.get_mut(replace_idx) {
        *slot = warning;
    } else {
        // Should be unreachable, but fall back to the bounded helper for safety.
        push_warning_bounded(warnings, warning);
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
    /// Record offset for the first `FILTERMODE` record seen in this worksheet substream (when any).
    pub(crate) filter_mode_offset: Option<usize>,
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

    let mut scanned = 0usize;
    for record in iter {
        let record = match record {
            Ok(r) => r,
            Err(err) => {
                push_sheet_metadata_warning(
                    &mut out.warnings,
                    format!("malformed BIFF record: {err}"),
                );
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_sheet_metadata_warning_force(
                    &mut out.warnings,
                    "record counter overflow while scanning sheet protection; stopping early",
                );
                break;
            }
        };
        if scanned > MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN {
            push_sheet_metadata_warning_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning sheet protection (cap={MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN}); stopping early"
                ),
            );
            break;
        }

        let data = record.data.as_ref();
        match record.record_id {
            RECORD_PROTECT => {
                if data.len() < 2 {
                    push_sheet_metadata_warning(
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
                    push_sheet_metadata_warning(
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
                    push_sheet_metadata_warning(
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
                    push_sheet_metadata_warning(
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
                    Err(err) => push_sheet_metadata_warning(
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
                    Err(err) => push_sheet_metadata_warning(
                        &mut out.warnings,
                        format!(
                            "failed to parse FEAT record at offset {}: {err}",
                            record.offset
                        ),
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

// Bits we understand (and are able to map into `formula_model::SheetProtection`) within the
// enhanced protection allow-mask.
//
// Some BIFF writers appear to include additional bits in the same mask (or store the mask inside a
// larger structure). We treat unknown bits as "don't care" and only act on these known flags.
const KNOWN_ALLOW_MASK_BITS: u16 = (ALLOW_SELECT_LOCKED_CELLS
    | ALLOW_SELECT_UNLOCKED_CELLS
    | ALLOW_FORMAT_CELLS
    | ALLOW_FORMAT_COLUMNS
    | ALLOW_FORMAT_ROWS
    | ALLOW_INSERT_COLUMNS
    | ALLOW_INSERT_ROWS
    | ALLOW_INSERT_HYPERLINKS
    | ALLOW_DELETE_COLUMNS
    | ALLOW_DELETE_ROWS
    | ALLOW_SORT
    | ALLOW_AUTO_FILTER
    | ALLOW_PIVOT_TABLES
    // Some writers use bits 14/15 for selection flags.
    | (1u32 << 14)
    | (1u32 << 15)) as u16;

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

    let isf_end = base
        .checked_add(2)
        .ok_or_else(|| "FEAT record offset overflow while reading isf".to_string())?;
    let isf_bytes = data
        .get(base..isf_end)
        .ok_or_else(|| "FEAT record truncated while reading isf".to_string())?;
    let isf = u16::from_le_bytes([isf_bytes[0], isf_bytes[1]]);
    if isf != FEAT_ISF_SHEET_PROTECTION {
        return Ok(None);
    }

    let cb_start = base
        .checked_add(4)
        .ok_or_else(|| "FEAT record offset overflow while reading cbFeatData".to_string())?;
    let cb_end = cb_start
        .checked_add(4)
        .ok_or_else(|| "FEAT record offset overflow while reading cbFeatData".to_string())?;
    let cb_feat_data_bytes = data
        .get(cb_start..cb_end)
        .ok_or_else(|| "FEAT record truncated while reading cbFeatData".to_string())?;
    let cb_feat_data = u32::from_le_bytes([
        cb_feat_data_bytes[0],
        cb_feat_data_bytes[1],
        cb_feat_data_bytes[2],
        cb_feat_data_bytes[3],
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
    let feat_data = data.get(data_start..data_end).ok_or_else(|| {
        debug_assert!(
            false,
            "FEAT payload slice out of bounds (len={}, data_start={data_start}, data_end={data_end})",
            data.len()
        );
        "FEAT payload slice out of bounds".to_string()
    })?;
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

    let isf_end = base
        .checked_add(2)
        .ok_or_else(|| "FEATHEADR record offset overflow while reading isf".to_string())?;
    let isf_bytes = data
        .get(base..isf_end)
        .ok_or_else(|| "FEATHEADR record truncated while reading isf".to_string())?;
    let isf = u16::from_le_bytes([isf_bytes[0], isf_bytes[1]]);
    if isf != FEAT_ISF_SHEET_PROTECTION {
        return Ok(None);
    }

    let cb_start = base
        .checked_add(4)
        .ok_or_else(|| "FEATHEADR record offset overflow while reading cbHdrData".to_string())?;
    let cb_end = cb_start
        .checked_add(4)
        .ok_or_else(|| "FEATHEADR record offset overflow while reading cbHdrData".to_string())?;
    let cb_hdr_data_bytes = data
        .get(cb_start..cb_end)
        .ok_or_else(|| "FEATHEADR record truncated while reading cbHdrData".to_string())?;
    let cb_hdr_data = u32::from_le_bytes([
        cb_hdr_data_bytes[0],
        cb_hdr_data_bytes[1],
        cb_hdr_data_bytes[2],
        cb_hdr_data_bytes[3],
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
    let hdr_data = data.get(data_start..data_end).ok_or_else(|| {
        debug_assert!(
            false,
            "FEATHEADR payload slice out of bounds (len={}, data_start={data_start}, data_end={data_end})",
            data.len()
        );
        "FEATHEADR payload slice out of bounds".to_string()
    })?;
    let mask = parse_allow_mask_best_effort(hdr_data)
        .ok_or_else(|| "FEATHEADR protection payload missing allow-mask".to_string())?;
    Ok(Some(mask))
}

fn parse_allow_mask_best_effort(payload: &[u8]) -> Option<u32> {
    // Best-effort: some writers store the allow-mask as a u16 at the start of the structure, while
    // others embed it deeper in the FEAT/FEATHEADR payload. Scan for a plausible u16 mask (one that
    // only uses known bits) and prefer the candidate with the most bits set.
    if payload.len() < 2 {
        return None;
    }

    let mut best: Option<(u32, usize, u16)> = None;
    for offset in 0..=(payload.len() - 2) {
        let Some(end) = offset.checked_add(2) else {
            break;
        };
        let Some(bytes) = payload.get(offset..end) else {
            break;
        };
        let mask = u16::from_le_bytes([bytes[0], bytes[1]]);
        if (mask & !KNOWN_ALLOW_MASK_BITS) != 0 {
            continue;
        }
        let score = mask.count_ones();
        match best {
            None => best = Some((score, offset, mask)),
            Some((best_score, best_offset, _)) => {
                if score > best_score || (score == best_score && offset < best_offset) {
                    best = Some((score, offset, mask));
                }
            }
        }
    }

    if let Some((_score, _offset, mask)) = best {
        return Some(mask as u32);
    }

    // Fall back to the first u16 even if it contains unknown bits; this preserves the prior
    // behavior for files we don't fully understand.
    payload
        .get(0..2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) as u32)
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
    let mut warned_selection_records_capped = false;
    let mut scanned = 0usize;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(err) => {
                push_sheet_metadata_warning(
                    &mut out.warnings,
                    format!("malformed BIFF record: {err}"),
                );
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_sheet_metadata_warning_force(
                    &mut out.warnings,
                    "record counter overflow while scanning sheet view state; stopping early",
                );
                break;
            }
        };
        if scanned > MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN {
            push_sheet_metadata_warning_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning sheet view state (cap={MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN}); stopping early"
                ),
            );
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
                Err(err) => push_sheet_metadata_warning(
                    &mut out.warnings,
                    format!("failed to parse WINDOW2 record: {err}"),
                ),
            },
            RECORD_SCL => match parse_scl_zoom(data) {
                Ok(zoom) => out.zoom = Some(zoom),
                Err(err) => push_sheet_metadata_warning(
                    &mut out.warnings,
                    format!("failed to parse SCL record: {err}"),
                ),
            },
            RECORD_PANE => match parse_pane_record(data, window2_frozen) {
                Ok((pane, pnn_act)) => {
                    out.pane = Some(pane);
                    active_pane = Some(pnn_act);
                }
                Err(err) => push_sheet_metadata_warning(
                    &mut out.warnings,
                    format!("failed to parse PANE record: {err}"),
                ),
            },
            RECORD_SELECTION => match parse_selection_record_best_effort(data) {
                Ok((pane, selection, summary)) => {
                    if summary.declared_refs > summary.parsed_refs {
                        push_sheet_metadata_warning_force(
                            &mut out.warnings,
                            format!(
                                "SELECTION record at offset {} declares cref={} refs; parsed {} (available {}, cap={})",
                                record.offset,
                                summary.declared_refs,
                                summary.parsed_refs,
                                summary.available_refs,
                                MAX_SELECTION_RANGES_PER_RECORD
                            ),
                        );
                    }

                    // Deduplicate by pane id; for well-formed files we expect at most one selection
                    // per pane. Stop retaining new pane ids once we hit the cap.
                    if let Some(existing) = selections.iter_mut().find(|(p, _)| *p == pane) {
                        existing.1 = selection;
                    } else if selections.len() < MAX_SELECTION_RECORDS_PER_SHEET_VIEW_STATE {
                        selections.push((pane, selection));
                    } else if !warned_selection_records_capped {
                        push_sheet_metadata_warning_force(
                            &mut out.warnings,
                            format!(
                                "too many SELECTION records (cap={MAX_SELECTION_RECORDS_PER_SHEET_VIEW_STATE}); additional selections ignored"
                            ),
                        );
                        warned_selection_records_capped = true;
                    }
                }
                Err(err) => push_sheet_metadata_warning(
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
/// break occurs; we therefore subtract 1 when importing.
///
/// Some producers emit `row=0` / `col=0`, which would represent a break *before* the first row/col.
/// Since `ManualPageBreaks` stores indices **after which** a break occurs, such breaks are not
/// representable and are ignored (with a warning).
///
/// This scan is resilient to malformed records: payload-level parse failures are surfaced as
/// warnings and otherwise ignored.
pub(crate) fn parse_biff_sheet_manual_page_breaks(
    workbook_stream: &[u8],
    start: usize,
) -> Result<BiffSheetManualPageBreaks, String> {
    let mut out = BiffSheetManualPageBreaks::default();

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, start)?;
    let mut scanned = 0usize;

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

        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_warning_bounded_force(
                    &mut out.warnings,
                    "record counter overflow while scanning sheet manual page breaks; stopping early",
                );
                break;
            }
        };
        if scanned > MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN {
            push_warning_bounded_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning sheet manual page breaks (cap={MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN}); stopping early"
                ),
            );
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
    // [MS-XLS] 2.4.122 HORIZONTALPAGEBREAKS caps `cbrk` at 1026.
    const SPEC_MAX: usize = 1026;
    const ENTRY_LEN: usize = 6;

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

    let max_entries_by_len = data.len().checked_sub(2).unwrap_or(0) / ENTRY_LEN;
    let iter_entries = cbrk.min(max_entries_by_len).min(SPEC_MAX);

    if cbrk > iter_entries {
        push_warning_bounded_force(
            warnings,
            format!(
                "HorizontalPageBreaks record at offset {record_offset}: cbrk={cbrk} exceeds available entries (payload_len={}, max_entries_by_len={max_entries_by_len}, spec_max={SPEC_MAX}); parsing {iter_entries} entries",
                data.len()
            ),
        );
    }

    for i in 0..iter_entries {
        let Some(base) = i
            .checked_mul(ENTRY_LEN)
            .and_then(|n| n.checked_add(2usize))
        else {
            push_warning_bounded(
                warnings,
                format!(
                    "HorizontalPageBreaks record at offset {record_offset}: overflow computing entry offset for entry {i}"
                ),
            );
            break;
        };
        let Some(bytes) = data.get(base..).and_then(|rest| rest.get(..2)) else {
            push_warning_bounded(
                warnings,
                format!(
                    "HorizontalPageBreaks record at offset {record_offset} truncated while reading entry {i}"
                ),
            );
            break;
        };
        let row = u16::from_le_bytes([bytes[0], bytes[1]]);
        if row == 0 {
            // `row=0` would represent a break before the first row, which is not representable in
            // `ManualPageBreaks` (it stores indices *after which* a break occurs).
            push_warning_bounded(warnings, "ignoring horizontal page break with row=0");
            continue;
        }
        manual_page_breaks.row_breaks_after.insert((row - 1) as u32);
    }
}

fn parse_vertical_page_breaks_record(
    data: &[u8],
    record_offset: usize,
    manual_page_breaks: &mut ManualPageBreaks,
    warnings: &mut Vec<String>,
) {
    // [MS-XLS] 2.4.350 VERTICALPAGEBREAKS caps `cbrk` at 255.
    const SPEC_MAX: usize = 255;
    const ENTRY_LEN: usize = 6;

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

    let max_entries_by_len = data.len().checked_sub(2).unwrap_or(0) / ENTRY_LEN;
    let iter_entries = cbrk.min(max_entries_by_len).min(SPEC_MAX);

    if cbrk > iter_entries {
        push_warning_bounded_force(
            warnings,
            format!(
                "VerticalPageBreaks record at offset {record_offset}: cbrk={cbrk} exceeds available entries (payload_len={}, max_entries_by_len={max_entries_by_len}, spec_max={SPEC_MAX}); parsing {iter_entries} entries",
                data.len()
            ),
        );
    }

    for i in 0..iter_entries {
        let Some(base) = i
            .checked_mul(ENTRY_LEN)
            .and_then(|n| n.checked_add(2usize))
        else {
            push_warning_bounded(
                warnings,
                format!(
                    "VerticalPageBreaks record at offset {record_offset}: overflow computing entry offset for entry {i}"
                ),
            );
            break;
        };
        let Some(bytes) = data.get(base..).and_then(|rest| rest.get(..2)) else {
            push_warning_bounded(
                warnings,
                format!(
                    "VerticalPageBreaks record at offset {record_offset} truncated while reading entry {i}"
                ),
            );
            break;
        };
        let col = u16::from_le_bytes([bytes[0], bytes[1]]);
        if col == 0 {
            // `col=0` would represent a break before the first column, which is not representable
            // in `ManualPageBreaks` (it stores indices *after which* a break occurs).
            push_warning_bounded(warnings, "ignoring vertical page break with col=0");
            continue;
        }
        manual_page_breaks.col_breaks_after.insert((col - 1) as u32);
    }
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

fn parse_selection_record_best_effort(
    data: &[u8],
) -> Result<(u16, SheetSelection, SelectionRecordSummary), String> {
    // Try a small set of plausible BIFF8 layouts.
    //
    // Different producers (and BIFF versions) vary in whether the pane id is stored as u8 vs u16
    // and whether the selection refs use the 6-byte RefU vs 8-byte Ref8 encoding.
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU8NoPadRefU) {
        return Ok(v);
    }
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU8PadRefU) {
        return Ok(v);
    }
    if let Ok(v) = parse_selection_record(data, SelectionLayout::PnnU16Ref8) {
        return Ok(v);
    }

    Err("unrecognized SELECTION record layout".to_string())
}

#[derive(Debug, Clone, Copy)]
enum SelectionLayout {
    PnnU8NoPadRefU,
    PnnU8PadRefU,
    PnnU16Ref8,
}

#[derive(Debug, Clone, Copy)]
struct SelectionRecordSummary {
    declared_refs: usize,
    parsed_refs: usize,
    available_refs: usize,
}

fn parse_selection_record(
    data: &[u8],
    layout: SelectionLayout,
) -> Result<(u16, SheetSelection, SelectionRecordSummary), String> {
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
    let payload_len = data
        .len()
        .checked_sub(refs_start)
        .ok_or_else(|| "SELECTION refs start out of bounds".to_string())?;
    let available_refs = payload_len / ref_len;
    let parsed_refs = cref_usize
        .min(available_refs)
        .min(MAX_SELECTION_RANGES_PER_RECORD);

    let active_row_u32 = rw_active as u32;
    let active_col_u32 = col_active as u32;
    if active_row_u32 >= EXCEL_MAX_ROWS || active_col_u32 >= EXCEL_MAX_COLS {
        return Err(format!(
            "active cell out of bounds: row={active_row_u32} col={active_col_u32}"
        ));
    }
    let active_cell = CellRef::new(active_row_u32, active_col_u32);

    let mut ranges = Vec::new();
    let _ = ranges.try_reserve_exact(parsed_refs);
    let mut off = refs_start;
    for _ in 0..parsed_refs {
        let range = match layout {
            SelectionLayout::PnnU16Ref8 => {
                let Some(end) = off.checked_add(8) else {
                    return Err("SELECTION refs offset overflow while reading Ref8 ranges".to_string());
                };
                let Some(chunk) = data.get(off..end) else {
                    debug_assert!(
                        false,
                        "SELECTION refs out of bounds (off={off}, len={})",
                        data.len()
                    );
                    return Err("SELECTION record truncated while reading Ref8 ranges".to_string());
                };
                let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]) as u32;
                let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]) as u32;
                let col_first = u16::from_le_bytes([chunk[4], chunk[5]]) as u32;
                let col_last = u16::from_le_bytes([chunk[6], chunk[7]]) as u32;
                off = end;
                make_range(rw_first, rw_last, col_first, col_last)?
            }
            SelectionLayout::PnnU8NoPadRefU | SelectionLayout::PnnU8PadRefU => {
                let Some(end) = off.checked_add(6) else {
                    return Err("SELECTION refs offset overflow while reading RefU ranges".to_string());
                };
                let Some(chunk) = data.get(off..end) else {
                    debug_assert!(
                        false,
                        "SELECTION refs out of bounds (off={off}, len={})",
                        data.len()
                    );
                    return Err("SELECTION record truncated while reading RefU ranges".to_string());
                };
                let rw_first = u16::from_le_bytes([chunk[0], chunk[1]]) as u32;
                let rw_last = u16::from_le_bytes([chunk[2], chunk[3]]) as u32;
                let col_first = chunk[4] as u32;
                let col_last = chunk[5] as u32;
                off = end;
                make_range(rw_first, rw_last, col_first, col_last)?
            }
        };
        ranges.push(range);
    }

    Ok((
        pane,
        SheetSelection::new(active_cell, ranges),
        SelectionRecordSummary {
            declared_refs: cref_usize,
            parsed_refs,
            available_refs,
        },
    ))
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
    let mut pending_autofilter12: Option<PendingFrtPayload> = None;

    let mut saw_eof = false;
    let mut warned_colinfo_first_oob = false;
    // Some worksheet-level records (SORT and BIFF8 Future Record Type records like AutoFilter12)
    // may legally be split across `CONTINUE` records. Use the logical iterator so we can
    // reassemble those payloads before decoding.
    let allows_continuation =
        |record_id: u16| record_id == RECORD_SORT || (record_id >= 0x0850 && record_id <= 0x08FF);
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;
    for record in iter {
        let record = match record {
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
            flush_pending_autofilter12_record(
                pending_autofilter12.take(),
                codepage,
                &mut autofilter12_columns,
                &mut props,
            );
            break;
        }

        // Flush any pending AutoFilter12 record as soon as we encounter a non-continuation record.
        // This keeps record association deterministic and ensures we don't accidentally attach
        // continuation bytes to the wrong record type.
        if record.record_id != RECORD_CONTINUEFRT12 {
            flush_pending_autofilter12_record(
                pending_autofilter12.take(),
                codepage,
                &mut autofilter12_columns,
                &mut props,
            );
        }

        let data = record.data.as_ref();
        match record.record_id {
            // DIMENSIONS [MS-XLS 2.4.84]
            RECORD_DIMENSIONS => {
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
                props.filter_mode_offset.get_or_insert(record.offset);
            }
            // SORT [MS-XLS 2.4.256]
            RECORD_SORT => match parse_sort_record_best_effort(data) {
                Ok(Some(state)) => {
                    // Prefer the last SORT record in the sheet stream (Excel may emit multiple
                    // records as sort state evolves).
                    props.sort_state = Some(state);
                }
                Ok(None) => {}
                Err(err) => push_warning_bounded(
                    &mut props.warnings,
                    format!(
                        "failed to parse SORT record at offset {}: {err}",
                        record.offset
                    ),
                ),
            },
            // AutoFilter12 / Sort12 / SortData12 (BIFF8 Future Record Type records).
            //
            // These records start with an `FrtHeader` structure. The record id in the BIFF
            // stream is often the same as `FrtHeader.rt`, but we still key off `rt` for
            // robustness.
            RECORD_CONTINUEFRT12 => {
                // Best-effort continuation for AutoFilter12 payloads.
                let Some(mut pending) = pending_autofilter12.take() else {
                    continue;
                };
                let header_len = if data.len() >= 8 { 8 } else { 0 };
                let fragments: Vec<&[u8]> = record.fragments().collect();

                let mut remaining_header = header_len;
                let mut dropped = false;
                for frag in fragments {
                    let payload = if remaining_header > 0 {
                        if frag.len() <= remaining_header {
                            remaining_header -= frag.len();
                            continue;
                        }
                        let start = remaining_header;
                        remaining_header = 0;
                        &frag[start..]
                    } else {
                        frag
                    };

                    if payload.is_empty() {
                        continue;
                    }

                    if pending.fragment_sizes.len() >= records::MAX_LOGICAL_RECORD_FRAGMENTS {
                        push_warning_bounded_force(
                            &mut props.warnings,
                            format!(
                                "too many ContinueFrt12 fragments (cap={}); dropping continued AutoFilter12",
                                records::MAX_LOGICAL_RECORD_FRAGMENTS
                            ),
                        );
                        dropped = true;
                        break;
                    }
                    let Some(next_len) = pending.payload.len().checked_add(payload.len()) else {
                        push_warning_bounded_force(
                            &mut props.warnings,
                            format!(
                                "AutoFilter12 continued payload length overflow (cap={} bytes); dropping continued AutoFilter12",
                                records::MAX_LOGICAL_RECORD_BYTES
                            ),
                        );
                        dropped = true;
                        break;
                    };
                    if next_len > records::MAX_LOGICAL_RECORD_BYTES {
                        push_warning_bounded_force(
                            &mut props.warnings,
                            format!(
                                "AutoFilter12 continued payload too large (cap={} bytes); dropping continued AutoFilter12",
                                records::MAX_LOGICAL_RECORD_BYTES
                            ),
                        );
                        dropped = true;
                        break;
                    }

                    pending.payload.extend_from_slice(payload);
                    pending.fragment_sizes.push(payload.len());
                }

                if !dropped {
                    pending_autofilter12 = Some(pending);
                }
            }
            id if id >= 0x0850 && id <= 0x08FF => {
                let Some((rt, _)) = parse_frt_header(data) else {
                    // Not a valid FRT header; ignore silently.
                    continue;
                };

                match rt {
                    RECORD_AUTOFILTER12 => {
                        saw_autofilter12 = true;
                        // AutoFilter12 records may be continued via one or more ContinueFrt12
                        // records. Stash the payload and decode once any continuations are seen.
                        let fragments: Vec<&[u8]> = record.fragments().collect();
                        let mut pending = PendingFrtPayload {
                            payload: Vec::new(),
                            fragment_sizes: Vec::new(),
                        };
                        // FrtHeader is 8 bytes.
                        let mut remaining_header = 8usize;
                        for frag in fragments {
                            let payload = if remaining_header > 0 {
                                if frag.len() <= remaining_header {
                                    remaining_header -= frag.len();
                                    continue;
                                }
                                let start = remaining_header;
                                remaining_header = 0;
                                &frag[start..]
                            } else {
                                frag
                            };

                            if payload.is_empty() {
                                continue;
                            }
                            // Avoid holding on to pathological payloads in tests where
                            // MAX_LOGICAL_RECORD_BYTES is small.
                            let Some(next_len) = pending.payload.len().checked_add(payload.len())
                            else {
                                push_warning_bounded_force(
                                    &mut props.warnings,
                                    format!(
                                        "AutoFilter12 payload length overflow (cap={} bytes); dropping AutoFilter12",
                                        records::MAX_LOGICAL_RECORD_BYTES
                                    ),
                                );
                                pending.payload.clear();
                                pending.fragment_sizes.clear();
                                break;
                            };
                            if next_len > records::MAX_LOGICAL_RECORD_BYTES {
                                push_warning_bounded_force(
                                    &mut props.warnings,
                                    format!(
                                        "AutoFilter12 payload too large (cap={} bytes); dropping AutoFilter12",
                                        records::MAX_LOGICAL_RECORD_BYTES
                                    ),
                                );
                                pending.payload.clear();
                                pending.fragment_sizes.clear();
                                break;
                            }
                            pending.payload.extend_from_slice(payload);
                            pending.fragment_sizes.push(payload.len());
                        }

                        if !pending.payload.is_empty() {
                            pending_autofilter12 = Some(pending);
                        }
                    }
                    // Sort12/SortData12 future records are imported separately during AutoFilter
                    // post-processing (see `biff::sort`).
                    RECORD_SORT12
                    | RECORD_SORT12_ALT
                    | RECORD_SORTDATA12
                    | RECORD_SORTDATA12_ALT => {}
                    _ => {}
                }
            }
            // ROW [MS-XLS 2.4.184]
            RECORD_ROW => {
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

                    let max_col = EXCEL_MAX_COLS - 1;
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

    // Flush any trailing AutoFilter12 payload if the sheet stream ended without another record
    // boundary (e.g. truncated sheet missing EOF).
    flush_pending_autofilter12_record(
        pending_autofilter12.take(),
        codepage,
        &mut autofilter12_columns,
        &mut props,
    );

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
                let mut end_row = last_row_plus1 - 1;
                let mut end_col = last_col_plus1 - 1;

                if first_row >= EXCEL_MAX_ROWS || first_col >= EXCEL_MAX_COLS {
                    // Ignore out-of-bounds dimensions.
                } else {
                    end_row = end_row.min(EXCEL_MAX_ROWS - 1);
                    end_col = end_col.min(EXCEL_MAX_COLS - 1);

                    if let Some(cols) = autofilter_cols {
                        if cols > 0 {
                            if let Some(last_filter_col) = cols
                                .checked_sub(1)
                                .and_then(|d| first_col.checked_add(d))
                            {
                                end_col = end_col.min(last_filter_col);
                            }
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

#[derive(Debug)]
struct PendingFrtPayload {
    payload: Vec<u8>,
    fragment_sizes: Vec<usize>,
}

fn flush_pending_autofilter12_record(
    pending: Option<PendingFrtPayload>,
    codepage: u16,
    columns: &mut BTreeMap<u32, FilterColumn>,
    props: &mut SheetRowColProperties,
) {
    let Some(pending) = pending else {
        return;
    };

    match decode_autofilter12_record(&pending.payload, &pending.fragment_sizes, codepage) {
        Ok(Some(column)) => {
            columns.entry(column.col_id).or_insert(column);
        }
        Ok(None) | Err(_) => {
            if !props
                .warnings
                .iter()
                .any(|w| w == "unsupported AutoFilter12")
            {
                push_warning_bounded(&mut props.warnings, "unsupported AutoFilter12".to_string());
            }
        }
    }
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
        start_row += 1;
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

// BIFF8 string flags used by `XLUnicodeString`.
// See [MS-XLS] 2.5.268.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

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
    fragment_sizes: &[usize],
    codepage: u16,
) -> Result<Option<FilterColumn>, String> {
    fn locate_fragment_offset(
        fragment_sizes: &[usize],
        global_offset: usize,
    ) -> Option<(usize, usize)> {
        let mut remaining = global_offset;
        for (idx, &size) in fragment_sizes.iter().enumerate() {
            if remaining < size {
                return Some((idx, remaining));
            }
            remaining -= size;
        }
        None
    }

    fn build_fragments<'a>(
        payload: &'a [u8],
        fragment_sizes: &[usize],
    ) -> Result<Vec<&'a [u8]>, String> {
        if fragment_sizes.is_empty() {
            return Ok(vec![payload]);
        }
        let mut out = Vec::new();
        out.try_reserve_exact(fragment_sizes.len())
            .map_err(|_| "allocation failed (AutoFilter12 fragments)".to_string())?;
        let mut offset = 0usize;
        for &size in fragment_sizes {
            let end = offset
                .checked_add(size)
                .ok_or_else(|| "AutoFilter12 fragment size overflow".to_string())?;
            let frag = payload
                .get(offset..end)
                .ok_or_else(|| "AutoFilter12 fragment sizes exceed payload length".to_string())?;
            out.push(frag);
            offset = end;
        }
        if offset != payload.len() {
            // Defensive: when fragment sizes don't match, fall back to treating the payload as a
            // single fragment so we still parse best-effort without panicking.
            return Ok(vec![payload]);
        }
        Ok(out)
    }

    #[derive(Debug, Clone)]
    struct FragmentCursor<'a> {
        fragments: &'a [&'a [u8]],
        frag_idx: usize,
        offset: usize,
    }

    impl<'a> FragmentCursor<'a> {
        fn new(fragments: &'a [&'a [u8]], frag_idx: usize, offset: usize) -> Self {
            Self {
                fragments,
                frag_idx,
                offset,
            }
        }

        fn remaining_in_fragment(&self) -> usize {
            self.fragments
                .get(self.frag_idx)
                .map(|f| f.len().checked_sub(self.offset).unwrap_or(0))
                .unwrap_or(0)
        }

        fn advance_fragment(&mut self) -> Result<(), String> {
            self.frag_idx = self
                .frag_idx
                .checked_add(1)
                .ok_or_else(|| "fragment index overflow".to_string())?;
            self.offset = 0;
            if self.frag_idx >= self.fragments.len() {
                return Err("unexpected end of record".to_string());
            }
            Ok(())
        }

        fn read_u8(&mut self) -> Result<u8, String> {
            loop {
                let frag = self
                    .fragments
                    .get(self.frag_idx)
                    .ok_or_else(|| "unexpected end of record".to_string())?;
                if self.offset < frag.len() {
                    let b = frag[self.offset];
                    self.offset += 1;
                    return Ok(b);
                }
                self.advance_fragment()?;
            }
        }

        fn read_u16_le(&mut self) -> Result<u16, String> {
            let lo = self.read_u8()?;
            let hi = self.read_u8()?;
            Ok(u16::from_le_bytes([lo, hi]))
        }

        fn read_exact_from_current(&mut self, n: usize) -> Result<&'a [u8], String> {
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            let end = self
                .offset
                .checked_add(n)
                .ok_or_else(|| "offset overflow".to_string())?;
            if end > frag.len() {
                return Err("unexpected end of record".to_string());
            }
            let out = &frag[self.offset..end];
            self.offset = end;
            Ok(out)
        }

        fn advance_fragment_in_biff8_string(
            &mut self,
            is_unicode: &mut bool,
        ) -> Result<(), String> {
            self.advance_fragment()?;
            // When a BIFF8 string spans a CONTINUE boundary, Excel inserts a 1-byte option flags
            // prefix at the start of the continued fragment. The only relevant bit is `fHighByte`
            // (unicode vs compressed).
            let cont_flags = self.read_u8()?;
            *is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
            Ok(())
        }

        fn read_biff8_string_bytes(
            &mut self,
            mut n: usize,
            is_unicode: &mut bool,
        ) -> Result<Vec<u8>, String> {
            // Read `n` canonical bytes from a BIFF8 continued string payload, skipping the 1-byte
            // continuation flags prefix that appears at the start of each continued fragment.
            let total = n;
            let mut out = Vec::new();
            out.try_reserve_exact(total)
                .map_err(|_| "allocation failed (AutoFilter12 string bytes)".to_string())?;
            while n > 0 {
                if self.remaining_in_fragment() == 0 {
                    self.advance_fragment_in_biff8_string(is_unicode)?;
                    continue;
                }
                let available = self.remaining_in_fragment();
                let take = n.min(available);
                let bytes = self.read_exact_from_current(take)?;
                out.extend_from_slice(bytes);
                n -= take;
            }
            Ok(out)
        }

        fn skip_biff8_string_bytes(
            &mut self,
            mut n: usize,
            is_unicode: &mut bool,
        ) -> Result<(), String> {
            // Skip `n` canonical bytes from a BIFF8 continued string payload, consuming any inserted
            // continuation flags bytes at fragment boundaries.
            while n > 0 {
                if self.remaining_in_fragment() == 0 {
                    self.advance_fragment_in_biff8_string(is_unicode)?;
                    continue;
                }
                let available = self.remaining_in_fragment();
                let take = n.min(available);
                self.offset += take;
                n -= take;
            }
            Ok(())
        }

        fn read_biff8_unicode_string(&mut self, codepage: u16) -> Result<String, String> {
            // XLUnicodeString [MS-XLS 2.5.268]
            let cch = self.read_u16_le()? as usize;
            let flags = self.read_u8()?;

            let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;

            let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
                let bytes = self.read_biff8_string_bytes(2, &mut is_unicode)?;
                u16::from_le_bytes([bytes[0], bytes[1]]) as usize
            } else {
                0
            };

            let ext_size = if flags & STR_FLAG_EXT != 0 {
                let bytes = self.read_biff8_string_bytes(4, &mut is_unicode)?;
                u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize
            } else {
                0
            };

            let mut remaining_chars = cch;
            let mut out = String::new();

            while remaining_chars > 0 {
                if self.remaining_in_fragment() == 0 {
                    // Continuing character bytes into a new CONTINUE fragment: first byte is
                    // option flags for the continued segment (fHighByte).
                    self.advance_fragment_in_biff8_string(&mut is_unicode)?;
                    continue;
                }

                let bytes_per_char = if is_unicode { 2 } else { 1 };
                let available_bytes = self.remaining_in_fragment();
                let available_chars = available_bytes / bytes_per_char;
                if available_chars == 0 {
                    return Err("string continuation split mid-character".to_string());
                }

                let take_chars = remaining_chars.min(available_chars);
                let take_bytes = take_chars * bytes_per_char;
                let bytes = self.read_exact_from_current(take_bytes)?;

                if is_unicode {
                    let mut u16s = Vec::new();
                    u16s.try_reserve_exact(take_chars)
                        .map_err(|_| "allocation failed (utf16 chunk)".to_string())?;
                    for chunk in bytes.chunks_exact(2) {
                        u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                    }
                    out.push_str(&String::from_utf16_lossy(&u16s));
                } else {
                    out.push_str(&strings::decode_ansi(codepage, bytes));
                }

                remaining_chars -= take_chars;
            }

            let richtext_bytes = richtext_runs
                .checked_mul(4)
                .ok_or_else(|| "rich text run count overflow".to_string())?;
            let extra_len = richtext_bytes
                .checked_add(ext_size)
                .ok_or_else(|| "string ext payload length overflow".to_string())?;
            self.skip_biff8_string_bytes(extra_len, &mut is_unicode)?;

            Ok(out)
        }
    }

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

    let fragments = build_fragments(payload, fragment_sizes)?;
    // Ensure the fragment size vector matches the actual fragment slice layout used for parsing.
    // If `build_fragments` fell back to a single fragment, preserve that behavior here.
    let fragment_sizes: Vec<usize> = fragments.iter().map(|f| f.len()).collect();

    for &(count_off, vals_off) in candidates {
        if payload.len() < vals_off {
            continue;
        }
        let count_end = count_off
            .checked_add(2)
            .ok_or_else(|| "AutoFilter12 count offset overflow".to_string())?;
        let count_bytes = payload.get(count_off..count_end).ok_or_else(|| {
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
        let min_bytes = count.checked_mul(3).unwrap_or(usize::MAX);
        let available = payload.len().checked_sub(vals_off).unwrap_or(0);
        if count > MAX_VALUES || min_bytes > available {
            continue;
        }

        let Some((frag_idx, frag_off)) = locate_fragment_offset(&fragment_sizes, vals_off) else {
            continue;
        };
        let mut cursor = FragmentCursor::new(&fragments, frag_idx, frag_off);
        let mut values: Vec<String> = Vec::new();
        let _ = values.try_reserve_exact(count.min(16));
        while count > 0 {
            let Ok(mut s) = cursor.read_biff8_unicode_string(codepage) else {
                values.clear();
                break;
            };
            if s.contains('\0') {
                s.retain(|ch| ch != '\0');
            }
            values.push(s);
            count -= 1;
        }

        if values.is_empty() {
            continue;
        }

        let mut criteria = Vec::new();
        let _ = criteria.try_reserve_exact(values.len());
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
) -> Result<SheetMergedCells, String> {
    let mut out = SheetMergedCells::default();

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
                    let Some(end) = pos.checked_add(8) else {
                        break;
                    };
                    let Some(chunk) = data.get(pos..end) else {
                        break;
                    };
                    pos = end;

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

                    out.ranges.push(Range::new(
                        CellRef::new(rw_first, col_first),
                        CellRef::new(rw_last, col_last),
                    ));

                    if out.ranges.len() == MAX_MERGED_RANGES_PER_SHEET {
                        out.warnings.push(format!(
                            "too many merged ranges (cap={MAX_MERGED_RANGES_PER_SHEET}); stopping after {} ranges",
                            out.ranges.len()
                        ));
                        return Ok(out);
                    }
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

#[derive(Debug, Default)]
pub(crate) struct SheetMergedCells {
    pub(crate) ranges: Vec<Range>,
    pub(crate) warnings: Vec<String>,
}

pub(crate) fn parse_biff_sheet_cell_xf_indices_filtered(
    workbook_stream: &[u8],
    start: usize,
    xf_is_interesting: Option<&[bool]>,
) -> Result<HashMap<CellRef, u16>, String> {
    parse_biff_sheet_cell_xf_indices_filtered_with_cap(
        workbook_stream,
        start,
        xf_is_interesting,
        MAX_CELL_XF_ENTRIES_PER_SHEET,
    )
}

fn parse_biff_sheet_cell_xf_indices_filtered_with_cap(
    workbook_stream: &[u8],
    start: usize,
    xf_is_interesting: Option<&[bool]>,
    max_entries: usize,
) -> Result<HashMap<CellRef, u16>, String> {
    let mut out = HashMap::new();

    let mut maybe_insert = |row: u32, col: u32, xf: u16| -> Result<(), String> {
        if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
            return Ok(());
        }
        let cell = CellRef::new(row, col);
        if let Some(mask) = xf_is_interesting {
            let idx = xf as usize;
            // Retain out-of-range XF indices so callers can surface an aggregated warning.
            if idx >= mask.len() {
                if out.len() >= max_entries && !out.contains_key(&cell) {
                    return Err("too many cell XF entries; refusing to allocate".to_string());
                }
                out.insert(cell, xf);
                return Ok(());
            }
            if !mask[idx] {
                return Ok(());
            }
        }
        if out.len() >= max_entries && !out.contains_key(&cell) {
            return Err("too many cell XF entries; refusing to allocate".to_string());
        }
        out.insert(cell, xf);
        Ok(())
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
                maybe_insert(row, col, xf)?;
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
                let rk_data = &data[4..data.len() - 2];
                for (idx, chunk) in rk_data.chunks_exact(6).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf)?;
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
                let xf_data = &data[4..data.len() - 2];
                for (idx, chunk) in xf_data.chunks_exact(2).enumerate() {
                    let col = match col_first.checked_add(idx as u32) {
                        Some(col) => col,
                        None => break,
                    };
                    if col > col_last {
                        break;
                    }
                    let xf = u16::from_le_bytes([chunk[0], chunk[1]]);
                    maybe_insert(row, col, xf)?;
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
pub(crate) struct SheetFormulas {
    pub(crate) formulas: HashMap<CellRef, String>,
    pub(crate) warnings: Vec<String>,
}

/// Best-effort decode of BIFF8 worksheet formulas (`FORMULA` records).
///
/// This scans the worksheet substream for `FORMULA` records and decodes their BIFF8 `rgce` token
/// streams into formula text (without a leading `=`).
///
/// Notes:
/// - This parser is intentionally best-effort: malformed records are skipped and surfaced as
///   warnings. Unsupported tokens are rendered as placeholders by the rgce decoder.
/// - Only BIFF8 `rgce` decoding is supported; callers should guard on BIFF version.
pub(crate) fn parse_biff8_sheet_formulas(
    workbook_stream: &[u8],
    start: usize,
    ctx: &rgce::RgceDecodeContext<'_>,
) -> Result<SheetFormulas, String> {
    let mut out = SheetFormulas::default();
    // Prefer the richer worksheet formula parser so we can resolve shared-formula indirections
    // (`PtgExp` via `SHRFMLA`/`ARRAY`) rather than always rendering `PtgExp` as `#UNKNOWN!`.
    let mut parsed = worksheet_formulas::parse_biff8_worksheet_formulas(workbook_stream, start)?;
    for w in parsed.warnings.drain(..) {
        push_warning_bounded(&mut out.warnings, w.message);
    }

    for cell in parsed.formula_cells.values() {
        let mut resolved_rgce: Option<Vec<u8>> = None;
        let mut resolved_rgcb: Option<Vec<u8>> = None;
        // Most BIFF8 `rgce` tokens are self-contained (row/col fields are absolute with separate
        // relative/absolute flags). Some tokens (notably `PtgRefN`/`PtgAreaN`) require a "base cell"
        // coordinate to interpret relative offsets. In those cases, the base cell is:
        // - the current cell for normal/shared formulas
        // - the *array anchor* cell for array formulas (Excel shows the same array formula text for
        //   every cell in the group, anchored at the top-left/base cell).
        let mut decode_base = rgce::CellCoord::new(cell.cell.row, cell.cell.col);

        // Resolve PtgExp/PtgTbl when possible.
        if matches!(cell.rgce.first().copied(), Some(0x01 | 0x02)) {
            let mut resolve_warnings: Vec<crate::ImportWarning> = Vec::new();
            let resolution = worksheet_formulas::resolve_ptgexp_or_ptgtbl_best_effort(
                &parsed,
                cell,
                &mut resolve_warnings,
            );
            for w in resolve_warnings {
                push_warning_bounded(
                    &mut out.warnings,
                    {
                        let mut msg = String::new();
                        msg.push_str("cell ");
                        formula_model::push_a1_cell_ref(
                            cell.cell.row,
                            cell.cell.col,
                            false,
                            false,
                            &mut msg,
                        );
                        msg.push_str(": ");
                        msg.push_str(&w.message);
                        msg
                    },
                );
            }

            match resolution {
                worksheet_formulas::PtgReferenceResolution::Shared { base } => {
                    if let Some(def) = parsed.shrfmla.get(&base) {
                        resolved_rgce = super::formulas::materialize_biff8_rgce_from_base(
                            &def.rgce, base, cell.cell,
                        );
                        if resolved_rgce.is_some() {
                            resolved_rgcb = Some(def.rgcb.clone());
                        }
                    }
                }
                worksheet_formulas::PtgReferenceResolution::Array { base } => {
                    if let Some(def) = parsed.array.get(&base) {
                        // ARRAY formulas are shared across the group and should not be materialized
                        // per-cell (unlike SHRFMLA shared formulas). Excel displays the same array
                        // formula text for every cell in the array range, anchored at `base`.
                        resolved_rgce = Some(def.rgce.clone());
                        resolved_rgcb = Some(def.rgcb.clone());
                        decode_base = rgce::CellCoord::new(base.row, base.col);
                    }
                }
                _ => {}
            }
        }

        let rgce_bytes = resolved_rgce.as_deref().unwrap_or(&cell.rgce);
        let rgcb_bytes = resolved_rgcb.as_deref().unwrap_or(&cell.rgcb);
        let decoded = rgce::decode_biff8_rgce_with_base_and_rgcb(
            rgce_bytes,
            rgcb_bytes,
            ctx,
            Some(decode_base),
        );
        for warning in decoded.warnings {
            push_warning_bounded(
                &mut out.warnings,
                {
                    let mut msg = String::new();
                    msg.push_str("cell ");
                    formula_model::push_a1_cell_ref(
                        cell.cell.row,
                        cell.cell.col,
                        false,
                        false,
                        &mut msg,
                    );
                    msg.push_str(": ");
                    msg.push_str(&warning);
                    msg
                },
            );
        }
        if !decoded.text.trim().is_empty() {
            out.formulas.insert(cell.cell, decoded.text);
        }
    }

    Ok(out)
}

#[derive(Debug, Default)]
pub(crate) struct SheetTableFormulas {
    pub(crate) formulas: HashMap<CellRef, String>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TableArg {
    Missing,
    Ref(CellRef),
    RefError,
}

impl TableArg {
    fn render(self) -> String {
        match self {
            TableArg::Missing => String::new(),
            TableArg::Ref(cell) => {
                let mut out = String::new();
                formula_model::push_a1_cell_ref(cell.row, cell.col, false, false, &mut out);
                out
            }
            TableArg::RefError => "#REF!".to_string(),
        }
    }
}

#[derive(Debug, Clone)]
struct TableRecordDecoded {
    row_input: TableArg,
    col_input: TableArg,
}

impl TableRecordDecoded {
    fn render_formula(&self) -> String {
        // Preserve Excel's missing-argument syntax to keep the formula parseable, mirroring how
        // BIFF `PtgMissArg` tokens are rendered elsewhere (`TABLE(A1,)`, `TABLE(,B1)`).
        let row = self.row_input.render();
        let col = self.col_input.render();
        format!("TABLE({row},{col})")
    }
}

/// Best-effort scan of a BIFF8 worksheet substream for What-If Analysis data-table formulas.
///
/// BIFF8 represents data tables using:
/// - `PtgTbl` (`0x02`) as the *entire* formula token stream for data-table result cells, and
/// - a `TABLE` record (`0x0236`) in the worksheet substream that provides row/column input-cell
///   references.
///
/// Most BIFF formula decoders do not surface these as a textual `TABLE(...)` formula because the
/// `TABLE` record lives outside the `rgce` stream. For `.xls` worksheet import we *do* have that
/// worksheet context, so we can synthesize a stable, parseable formula string.
pub(crate) fn parse_biff8_sheet_table_formulas(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetTableFormulas, String> {
    // [MS-XLS] 2.5.198.21 (PtgTbl): [ptg:0x02][rw:u16][col:u16]
    const PTG_TBL: u8 = 0x02;

    // Best-effort guess for TABLE record flags indicating presence of input cells.
    //
    // The MS-XLS spec defines a `grbit` field, but producers in the wild vary. We interpret the
    // low bits as "row input present" and "col input present" when set, but if neither bit is set
    // we assume *both* inputs are present to avoid rendering `TABLE(,)` for two-input tables when
    // the flag semantics differ.
    const TABLE_GRBIT_HAS_ROW_INPUT: u16 = 0x0001;
    const TABLE_GRBIT_HAS_COL_INPUT: u16 = 0x0002;

    let mut out = SheetTableFormulas::default();

    let mut tables: HashMap<(u16, u16), TableRecordDecoded> = HashMap::new();

    #[derive(Debug, Clone)]
    enum PtgTblBase {
        Canonical { row: u16, col: u16 },
        WidePayload(Vec<u8>),
    }

    #[derive(Debug, Clone)]
    struct PendingPtgTbl {
        cell: CellRef,
        base: PtgTblBase,
        offset: usize,
    }

    let mut ptg_tbl_cells: Vec<PendingPtgTbl> = Vec::new();

    // FORMULA and TABLE records can legally be split across one or more `CONTINUE` records if the
    // payload exceeds the BIFF record size limit. Use the logical iterator so we can reassemble
    // those fragments before parsing PtgTbl references.
    let allows_continuation = |record_id: u16| matches!(record_id, RECORD_FORMULA | RECORD_TABLE);
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    for record in iter {
        let record = match record {
            Ok(r) => r,
            Err(err) => {
                push_warning_bounded(
                    &mut out.warnings,
                    format!("malformed BIFF record in sheet stream: {err}"),
                );
                break;
            }
        };

        // BOF indicates the start of a new substream; stop before yielding the next BOF in case
        // the worksheet is missing its EOF.
        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        let data = record.data.as_ref();
        match record.record_id {
            RECORD_TABLE => {
                // TABLE record payload (BIFF8) [MS-XLS 2.4.313].
                //
                // Common layout:
                //   [rw: u16][col: u16][grbit: u16]
                //   [rwInpRow: u16][colInpRow: u16]
                //   [rwInpCol: u16][colInpCol: u16]
                if data.len() < 4 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated TABLE record at offset {}: expected >=4 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }

                let base_row = u16::from_le_bytes([data[0], data[1]]);
                let base_col = u16::from_le_bytes([data[2], data[3]]) & 0x3FFF;

                let grbit = if data.len() >= 6 {
                    u16::from_le_bytes([data[4], data[5]])
                } else {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated TABLE record at offset {}: expected >=6 bytes for grbit, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    0
                };

                // Determine which input cells are present.
                let has_row_flag = (grbit & TABLE_GRBIT_HAS_ROW_INPUT) != 0;
                let has_col_flag = (grbit & TABLE_GRBIT_HAS_COL_INPUT) != 0;
                let assume_both_present = !has_row_flag && !has_col_flag;
                let row_present = has_row_flag || assume_both_present;
                let col_present = has_col_flag || assume_both_present;

                let decoded = if data.len() < 14 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated TABLE record at offset {}: expected >=14 bytes, got {}; rendering #REF! placeholders",
                            record.offset,
                            data.len()
                        ),
                    );

                    TableRecordDecoded {
                        row_input: if row_present {
                            TableArg::RefError
                        } else {
                            TableArg::Missing
                        },
                        col_input: if col_present {
                            TableArg::RefError
                        } else {
                            TableArg::Missing
                        },
                    }
                } else {
                    let row_inp_row = u16::from_le_bytes([data[6], data[7]]);
                    let row_inp_col_raw = u16::from_le_bytes([data[8], data[9]]);
                    let col_inp_row = u16::from_le_bytes([data[10], data[11]]);
                    let col_inp_col_raw = u16::from_le_bytes([data[12], data[13]]);

                    let decode_cell_ref = |rw: u16,
                                           col_raw: u16,
                                           label: &str,
                                           warnings: &mut Vec<String>|
                     -> TableArg {
                        let col = (col_raw & 0x3FFF) as u32;
                        let row = rw as u32;
                        if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
                            push_warning_bounded(
                                warnings,
                                format!(
                                    "TABLE record at offset {} has out-of-bounds {label} reference (row={row}, col={col}); rendering #REF!",
                                    record.offset
                                ),
                            );
                            return TableArg::RefError;
                        }
                        TableArg::Ref(CellRef::new(row, col))
                    };

                    let row_input = if row_present {
                        // Best-effort missing sentinel: some writers use 0xFFFF for an unused input.
                        if row_inp_col_raw == 0xFFFF {
                            TableArg::Missing
                        } else {
                            decode_cell_ref(
                                row_inp_row,
                                row_inp_col_raw,
                                "row input",
                                &mut out.warnings,
                            )
                        }
                    } else {
                        TableArg::Missing
                    };
                    let col_input = if col_present {
                        if col_inp_col_raw == 0xFFFF {
                            TableArg::Missing
                        } else {
                            decode_cell_ref(
                                col_inp_row,
                                col_inp_col_raw,
                                "col input",
                                &mut out.warnings,
                            )
                        }
                    } else {
                        TableArg::Missing
                    };

                    TableRecordDecoded {
                        row_input,
                        col_input,
                    }
                };

                if tables.insert((base_row, base_col), decoded).is_some() {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "duplicate TABLE record for base cell row={base_row} col={base_col} at offset {}; last one wins",
                            record.offset
                        ),
                    );
                }
            }
            RECORD_FORMULA => {
                // FORMULA record payload (BIFF8) [MS-XLS 2.4.127].
                if data.len() < 22 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated FORMULA record at offset {}: expected >=22 bytes, got {}",
                            record.offset,
                            data.len()
                        ),
                    );
                    continue;
                }

                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let col = u16::from_le_bytes([data[2], data[3]]) as u32;
                if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "skipping out-of-bounds FORMULA cell ({row},{col}) at offset {}",
                            record.offset
                        ),
                    );
                    continue;
                }
                let cell = CellRef::new(row, col);

                let cce = u16::from_le_bytes([data[20], data[21]]) as usize;
                let Some(rgce_end) = 22usize.checked_add(cce) else {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "overflow computing FORMULA rgce end at offset {} (cce={cce})",
                            record.offset
                        ),
                    );
                    continue;
                };
                let Some(rgce) = data.get(22..rgce_end) else {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated FORMULA rgce stream at offset {}: need {} bytes, got {}",
                            record.offset,
                            cce,
                            data.len().checked_sub(22).unwrap_or(0)
                        ),
                    );
                    continue;
                };

                if rgce.first().copied() != Some(PTG_TBL) {
                    continue;
                }

                if rgce.len() < 5 {
                    push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "truncated PtgTbl token in FORMULA record at offset {}: expected 4-byte payload, got {} bytes",
                            record.offset,
                            rgce.len().checked_sub(1).unwrap_or(0)
                        ),
                    );
                    out.formulas.insert(cell, "TABLE(#REF!,#REF!)".to_string());
                    continue;
                }

                if cce == 5 {
                    // Canonical BIFF8 payload: [ptg:0x02][rw:u16][col:u16]
                    let base_row = u16::from_le_bytes([rgce[1], rgce[2]]);
                    let base_col = u16::from_le_bytes([rgce[3], rgce[4]]) & 0x3FFF;
                    ptg_tbl_cells.push(PendingPtgTbl {
                        cell,
                        base: PtgTblBase::Canonical {
                            row: base_row,
                            col: base_col,
                        },
                        offset: record.offset,
                    });
                } else {
                    // Non-canonical payload width. Some writers emit BIFF12-style u32 row/col
                    // coordinates even in BIFF8 `.xls` files. Collect the raw payload bytes (after
                    // the PtgTbl opcode) and resolve later using `ptgexp_candidates`.
                    ptg_tbl_cells.push(PendingPtgTbl {
                        cell,
                        base: PtgTblBase::WidePayload(rgce[1..].to_vec()),
                        offset: record.offset,
                    });
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    // Resolve PtgTbl references after scanning the full substream so TABLE records can appear
    // before or after their referencing formula cells.
    for pending in ptg_tbl_cells {
        let (resolved_base, multiple_matches) = match pending.base {
            PtgTblBase::Canonical { row, col } => (Some((row, col)), false),
            PtgTblBase::WidePayload(payload) => {
                let candidates = worksheet_formulas::ptgexp_candidates(&payload);

                let mut matches: Vec<(u16, u16)> = Vec::new();
                for (r, c) in candidates {
                    let key = (r as u16, c as u16);
                    if tables.contains_key(&key) {
                        matches.push(key);
                    }
                }

                let resolved = matches.first().copied();
                (resolved, matches.len() > 1)
            }
        };

        let Some((base_row, base_col)) = resolved_base else {
            push_warning_bounded(
                &mut out.warnings,
                format!(
                    "PtgTbl formula at offset {} (cell {}) references missing TABLE record (no matching base cell candidates); rendering TABLE()",
                    pending.offset,
                    {
                        let mut a1 = String::new();
                        formula_model::push_a1_cell_ref(
                            pending.cell.row,
                            pending.cell.col,
                            false,
                            false,
                            &mut a1,
                        );
                        a1
                    }
                ),
            );
            out.formulas.insert(pending.cell, "TABLE()".to_string());
            continue;
        };

        if multiple_matches {
            push_warning_bounded(
                &mut out.warnings,
                format!(
                    "PtgTbl formula at offset {} (cell {}) has ambiguous wide payload; multiple TABLE record base-cell candidates matched; choosing row={base_row} col={base_col}",
                    pending.offset,
                    {
                        let mut a1 = String::new();
                        formula_model::push_a1_cell_ref(
                            pending.cell.row,
                            pending.cell.col,
                            false,
                            false,
                            &mut a1,
                        );
                        a1
                    }
                ),
            );
        }

        let formula = match tables.get(&(base_row, base_col)) {
            Some(def) => def.render_formula(),
            None => {
                push_warning_bounded(
                    &mut out.warnings,
                    format!(
                        "PtgTbl formula at offset {} (cell {}) references missing TABLE record at row={base_row} col={base_col}; rendering TABLE()",
                        pending.offset,
                        {
                            let mut a1 = String::new();
                            formula_model::push_a1_cell_ref(
                                pending.cell.row,
                                pending.cell.col,
                                false,
                                false,
                                &mut a1,
                            );
                            a1
                        }
                    ),
                );
                "TABLE()".to_string()
            }
        };

        out.formulas.insert(pending.cell, formula);
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

/// Hard cap on the number of BIFF records scanned while searching for hyperlinks.
///
/// The `.xls` importer performs multiple best-effort passes over each worksheet substream (e.g.
/// hyperlinks, notes, view state). Without a cap, a crafted workbook with millions of cell records
/// can force excessive work even when a particular feature is absent (e.g. a sheet with no
/// hyperlinks would still require scanning the entire substream).
#[cfg(not(test))]
const MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN: usize = 500_000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN: usize = 1_000;

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

    let mut scanned = 0usize;
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

        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_warning_bounded_force(
                    &mut out.warnings,
                    "record counter overflow while scanning sheet hyperlinks; stopping early",
                );
                break;
            }
        };
        if scanned > MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN {
            push_warning_bounded_force(
                &mut out.warnings,
                format!(
                    "too many BIFF records while scanning sheet hyperlinks (cap={MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN}); stopping early"
                ),
            );
            break;
        }

        match record.record_id {
            RECORD_HLINK => {
                // Hardening: once we hit the cap, stop scanning immediately to avoid additional
                // allocations and decode work on the rest of the sheet stream.
                if out.hyperlinks.len() >= MAX_HYPERLINKS_PER_SHEET {
                    push_warning_bounded_force(
                        &mut out.warnings,
                        "too many hyperlinks; additional HLINK records skipped".to_string(),
                    );
                    break;
                }

                match decode_hlink_record(record.data.as_ref(), codepage) {
                    Ok(Some(link)) => out.hyperlinks.push(link),
                    Ok(None) => {}
                    Err(err) => push_warning_bounded(
                        &mut out.warnings,
                        format!(
                            "failed to decode HLINK record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                }
            }
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

    let stream_version_end = pos
        .checked_add(4)
        .ok_or_else(|| "HLINK offset overflow while reading streamVersion".to_string())?;
    let stream_version_bytes = data
        .get(pos..stream_version_end)
        .ok_or_else(|| "HLINK record truncated while reading streamVersion".to_string())?;
    let stream_version = u32::from_le_bytes([
        stream_version_bytes[0],
        stream_version_bytes[1],
        stream_version_bytes[2],
        stream_version_bytes[3],
    ]);
    pos = stream_version_end;
    if stream_version != 2 {
        // Non-fatal; continue parsing.
        // Some producers may write a different version, but the layout is usually identical.
    }

    let link_opts_end = pos
        .checked_add(4)
        .ok_or_else(|| "HLINK offset overflow while reading linkOpts".to_string())?;
    let link_opts_bytes = data
        .get(pos..link_opts_end)
        .ok_or_else(|| "HLINK record truncated while reading linkOpts".to_string())?;
    let link_opts = u32::from_le_bytes([
        link_opts_bytes[0],
        link_opts_bytes[1],
        link_opts_bytes[2],
        link_opts_bytes[3],
    ]);
    pos = link_opts_end;

    let mut display: Option<String> = None;
    let mut tooltip: Option<String> = None;
    let mut text_mark: Option<String> = None;
    let mut uri: Option<String> = None;

    // Optional: display string.
    if (link_opts & HLINK_FLAG_HAS_DISPLAY) != 0 {
        let tail = data
            .get(pos..)
            .ok_or_else(|| "HLINK cursor out of bounds while reading display string".to_string())?;
        let (s, consumed) = parse_hyperlink_string(tail, codepage)?;
        display = (!s.is_empty()).then_some(s);
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: target frame (ignored for now).
    if (link_opts & HLINK_FLAG_HAS_TARGET_FRAME) != 0 {
        let tail = data.get(pos..).ok_or_else(|| {
            "HLINK cursor out of bounds while reading target frame".to_string()
        })?;
        let (_s, consumed) = parse_hyperlink_string(tail, codepage)?;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: moniker (external link target).
    if (link_opts & HLINK_FLAG_HAS_MONIKER) != 0 {
        let tail = data
            .get(pos..)
            .ok_or_else(|| "HLINK cursor out of bounds while reading moniker".to_string())?;
        let (parsed_uri, consumed) = parse_hyperlink_moniker(tail, codepage)?;
        uri = parsed_uri;
        pos = pos
            .checked_add(consumed)
            .ok_or_else(|| "HLINK offset overflow".to_string())?;
    }

    // Optional: location / text mark (internal target or sub-address).
    if (link_opts & HLINK_FLAG_HAS_LOCATION) != 0 {
        let tail = data
            .get(pos..)
            .ok_or_else(|| "HLINK cursor out of bounds while reading location".to_string())?;
        let (s, consumed) = parse_hyperlink_string(tail, codepage)?;
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

        if uri
            .get(.."mailto:".len())
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("mailto:"))
        {
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
    let clsid_bytes = input
        .get(..16)
        .ok_or_else(|| "truncated hyperlink moniker".to_string())?;
    let mut clsid = [0u8; 16];
    clsid.copy_from_slice(clsid_bytes);

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
            let Some(header) = input.get(pos..unicode_header_end) else {
                return Err("truncated file moniker unicode header".to_string());
            };
            let end_server = u16::from_le_bytes([header[0], header[1]]) as usize;
            // reserved/version (ignored)
            let _reserved = u16::from_le_bytes([header[2], header[3]]);
            let unicode_len =
                u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as usize;

            let available = input.len().checked_sub(unicode_header_end).unwrap_or(0);
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

    let mut out = String::new();
    let _ = out.try_reserve(path.len());
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    for &b in path.as_bytes() {
        if is_allowed(b) {
            out.push(b as char);
        } else {
            out.push('%');
            out.push(HEX[(b >> 4) as usize] as char);
            out.push(HEX[(b & 0x0F) as usize] as char);
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
    let p = if p.contains('\\') { p.replace('\\', "/") } else { p };
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

    #[derive(Clone, Copy)]
    struct Candidate {
        consumed: usize,
        ends_with_nul: bool,
    }

    let mut best: Option<Candidate> = None;

    let mut consider = |consumed: usize| {
        let bytes = &input[..consumed];
        let ends_with_nul = bytes
            .chunks_exact(2)
            .last()
            .is_some_and(|chunk| chunk[0] == 0 && chunk[1] == 0);
        let cand = Candidate {
            consumed,
            ends_with_nul,
        };
        best = match best {
            None => Some(cand),
            Some(prev) => {
                // Prefer NUL-terminated candidates; otherwise prefer the shorter byte length.
                let cand_key = (!cand.ends_with_nul, cand.consumed);
                let prev_key = (!prev.ends_with_nul, prev.consumed);
                if cand_key < prev_key {
                    Some(cand)
                } else {
                    Some(prev)
                }
            }
        };
    };

    // Candidate A: `len` as byte length.
    if len_as_bytes_ok && len % 2 == 0 && input.len() >= len {
        consider(len);
    }

    // Candidate B: `len` as character count.
    if let Some(byte_len) = len_as_chars_bytes {
        if len_as_chars_ok && byte_len % 2 == 0 && input.len() >= byte_len {
            consider(byte_len);
        }
    }

    let Some(best) = best else {
        return Err("truncated UTF-16 string".to_string());
    };

    let s = decode_utf16le(&input[..best.consumed])?;
    Ok((trim_at_first_nul(trim_trailing_nuls(s)), best.consumed))
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
                if let Some(end) = 4usize.checked_add(byte_len) {
                    if let Some(bytes) = input.get(4..end) {
                    let s = decode_utf16le(bytes)?;
                    // Hyperlink-related strings in BIFF are frequently NUL terminated, but we've
                    // observed files in the wild that include embedded NULs + trailing garbage
                    // within the declared length. Truncate at the first NUL for best-effort
                    // compatibility (mirrors how file moniker paths are handled).
                    let s = trim_at_first_nul(trim_trailing_nuls(s));
                    return Ok((s, end));
                    }
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
    let mut u16s = Vec::new();
    u16s.try_reserve_exact(bytes.len() / 2)
        .map_err(|_| "allocation failed (utf16 buffer)".to_string())?;
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
    let sheet = formula_model::unquote_sheet_name_lenient(sheet);

    let cell_str = cell.trim();
    let cell_str = cell_str
        .split_once(':')
        .map(|(start, _)| start)
        .unwrap_or(cell_str);
    let cell = CellRef::from_a1(cell_str).ok()?;
    Some((sheet, cell))
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_engine::{parse_formula, ParseOptions};
    use std::collections::BTreeSet;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
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
    fn utf16_prefixed_string_prefers_nul_terminated_candidate_over_shorter_candidate() {
        // `len=2` can be interpreted as 2 bytes ("A") or 2 UTF-16 code units ("A\\0"). Prefer the
        // NUL-terminated interpretation because it's common in monikers.
        let input = [0x41, 0x00, 0x00, 0x00];
        let (s, consumed) = parse_utf16_prefixed_string(&input, 2).expect("parse");
        assert_eq!(s, "A");
        assert_eq!(consumed, 4);
    }

    #[test]
    fn utf16_prefixed_string_falls_back_to_shorter_candidate_when_neither_is_nul_terminated() {
        // `len=4` can be interpreted as 4 bytes ("AB") or 4 UTF-16 code units ("ABCD"). When
        // neither candidate is NUL terminated, prefer the shorter one to avoid over-consuming.
        let input = [0x41, 0x00, 0x42, 0x00, 0x43, 0x00, 0x44, 0x00];
        let (s, consumed) = parse_utf16_prefixed_string(&input, 4).expect("parse");
        assert_eq!(s, "AB");
        assert_eq!(consumed, 4);
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
    fn percent_encode_uri_path_encodes_non_ascii_and_spaces_per_byte() {
        assert_eq!(percent_encode_uri_path("a b"), "a%20b");
        assert_eq!(percent_encode_uri_path("a/b:c"), "a/b:c");
        assert_eq!(percent_encode_uri_path("é"), "%C3%A9");
        assert_eq!(percent_encode_uri_path("💩"), "%F0%9F%92%A9");
    }

    #[test]
    fn decodes_ptgtbl_table_records_into_table_formula_text() {
        // TABLE record anchored at F11 (row=10, col=5), with two input cells: A1 (row input) and
        // B2 (col input).
        let base_row: u16 = 10;
        let base_col: u16 = 5;
        let grbit: u16 = 0x0003; // best-effort: both inputs present

        let mut table_payload = Vec::new();
        table_payload.extend_from_slice(&base_row.to_le_bytes());
        table_payload.extend_from_slice(&base_col.to_le_bytes());
        table_payload.extend_from_slice(&grbit.to_le_bytes());
        // row input: A1
        table_payload.extend_from_slice(&0u16.to_le_bytes()); // rwInpRow
        table_payload.extend_from_slice(&0u16.to_le_bytes()); // colInpRow
                                                              // col input: B2
        table_payload.extend_from_slice(&1u16.to_le_bytes()); // rwInpCol
        table_payload.extend_from_slice(&1u16.to_le_bytes()); // colInpCol

        // FORMULA record at D21 whose rgce is a single PtgTbl referencing the TABLE base cell.
        let cell_row: u16 = 20;
        let cell_col: u16 = 3;
        let rgce = [
            0x02u8, // PtgTbl
            base_row.to_le_bytes()[0],
            base_row.to_le_bytes()[1],
            base_col.to_le_bytes()[0],
            base_col.to_le_bytes()[1],
        ];

        let mut formula_payload = Vec::new();
        formula_payload.extend_from_slice(&cell_row.to_le_bytes());
        formula_payload.extend_from_slice(&cell_col.to_le_bytes());
        formula_payload.extend_from_slice(&0u16.to_le_bytes()); // xf
        formula_payload.extend_from_slice(&0f64.to_le_bytes()); // cached result
        formula_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        formula_payload.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_TABLE, &table_payload),
            record(RECORD_FORMULA, &formula_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_sheet_table_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let expected_cell = CellRef::new(cell_row as u32, cell_col as u32);
        assert_eq!(
            parsed.formulas.get(&expected_cell).map(String::as_str),
            Some("TABLE(A1,B2)")
        );

        // Ensure the synthesized formula is parseable by our formula lexer/parser.
        parse_formula("TABLE(A1,B2)", ParseOptions::default()).expect("parseable");
    }

    #[test]
    fn parses_table_formula_when_formula_record_is_continued() {
        // TABLE record anchored at F11 (row=10, col=5), with two input cells: A1 (row input) and
        // B2 (col input).
        let base_row: u16 = 10;
        let base_col: u16 = 5;
        let grbit: u16 = 0x0003; // best-effort: both inputs present

        let mut table_payload = Vec::new();
        table_payload.extend_from_slice(&base_row.to_le_bytes());
        table_payload.extend_from_slice(&base_col.to_le_bytes());
        table_payload.extend_from_slice(&grbit.to_le_bytes());
        // row input: A1
        table_payload.extend_from_slice(&0u16.to_le_bytes()); // rwInpRow
        table_payload.extend_from_slice(&0u16.to_le_bytes()); // colInpRow
        // col input: B2
        table_payload.extend_from_slice(&1u16.to_le_bytes()); // rwInpCol
        table_payload.extend_from_slice(&1u16.to_le_bytes()); // colInpCol

        // FORMULA record at D21 whose rgce is a single PtgTbl referencing the TABLE base cell.
        // Split across CONTINUE boundaries such that cce and rgce both cross fragment boundaries.
        let cell_row: u16 = 20;
        let cell_col: u16 = 3;
        let rgce = [
            0x02u8, // PtgTbl
            base_row.to_le_bytes()[0],
            base_row.to_le_bytes()[1],
            base_col.to_le_bytes()[0],
            base_col.to_le_bytes()[1],
        ];
        let cce_bytes = (rgce.len() as u16).to_le_bytes();

        // FORMULA header: row (2) + col (2) + xf (2) + cached result (8) + grbit (2) + chn (4) +
        // cce (2) + rgce...
        let mut formula_prefix = Vec::new();
        formula_prefix.extend_from_slice(&cell_row.to_le_bytes());
        formula_prefix.extend_from_slice(&cell_col.to_le_bytes());
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // xf
        formula_prefix.extend_from_slice(&0f64.to_le_bytes()); // cached result
        formula_prefix.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_prefix.extend_from_slice(&0u32.to_le_bytes()); // chn

        // Split so that the first CONTINUE boundary occurs after the first byte of cce (so cce
        // crosses), and the second CONTINUE boundary splits the rgce bytes.
        let formula_frag1 = [formula_prefix, vec![cce_bytes[0]]].concat();
        let cont1 = vec![cce_bytes[1], rgce[0], rgce[1]];
        let cont2 = rgce[2..].to_vec();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_TABLE, &table_payload),
            record(RECORD_FORMULA, &formula_frag1),
            record(records::RECORD_CONTINUE, &cont1),
            record(records::RECORD_CONTINUE, &cont2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_sheet_table_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let expected_cell = CellRef::new(cell_row as u32, cell_col as u32);
        assert_eq!(
            parsed.formulas.get(&expected_cell).map(String::as_str),
            Some("TABLE(A1,B2)")
        );

        parse_formula("TABLE(A1,B2)", ParseOptions::default()).expect("parseable");
    }

    #[test]
    fn parses_table_formula_when_table_record_is_continued() {
        // TABLE record anchored at F11 (row=10, col=5), with two input cells: A1 (row input) and
        // B2 (col input). Split across a CONTINUE boundary.
        let base_row: u16 = 10;
        let base_col: u16 = 5;
        let grbit: u16 = 0x0003; // best-effort: both inputs present

        // First fragment contains only base_row/base_col.
        let mut table_frag1 = Vec::new();
        table_frag1.extend_from_slice(&base_row.to_le_bytes());
        table_frag1.extend_from_slice(&base_col.to_le_bytes());

        // Second fragment contains grbit + input refs.
        let mut table_cont = Vec::new();
        table_cont.extend_from_slice(&grbit.to_le_bytes());
        // row input: A1
        table_cont.extend_from_slice(&0u16.to_le_bytes()); // rwInpRow
        table_cont.extend_from_slice(&0u16.to_le_bytes()); // colInpRow
        // col input: B2
        table_cont.extend_from_slice(&1u16.to_le_bytes()); // rwInpCol
        table_cont.extend_from_slice(&1u16.to_le_bytes()); // colInpCol

        // FORMULA record at D21 whose rgce is a single PtgTbl referencing the TABLE base cell.
        let cell_row: u16 = 20;
        let cell_col: u16 = 3;
        let rgce = [
            0x02u8, // PtgTbl
            base_row.to_le_bytes()[0],
            base_row.to_le_bytes()[1],
            base_col.to_le_bytes()[0],
            base_col.to_le_bytes()[1],
        ];

        let mut formula_payload = Vec::new();
        formula_payload.extend_from_slice(&cell_row.to_le_bytes());
        formula_payload.extend_from_slice(&cell_col.to_le_bytes());
        formula_payload.extend_from_slice(&0u16.to_le_bytes()); // xf
        formula_payload.extend_from_slice(&0f64.to_le_bytes()); // cached result
        formula_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        formula_payload.extend_from_slice(&0u32.to_le_bytes()); // chn
        formula_payload.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        formula_payload.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_TABLE, &table_frag1),
            record(records::RECORD_CONTINUE, &table_cont),
            record(RECORD_FORMULA, &formula_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_sheet_table_formulas(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );

        let expected_cell = CellRef::new(cell_row as u32, cell_col as u32);
        assert_eq!(
            parsed.formulas.get(&expected_cell).map(String::as_str),
            Some("TABLE(A1,B2)")
        );

        parse_formula("TABLE(A1,B2)", ParseOptions::default()).expect("parseable");
    }

    #[test]
    fn sheet_table_formulas_warnings_are_bounded() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Emit many malformed TABLE records (payload < 4 bytes). Without bounding this would
        // allocate unbounded warning strings.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 100) {
            stream.extend_from_slice(&record(RECORD_TABLE, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff8_sheet_table_formulas(&stream, 0).expect("parse");
        assert_eq!(parsed.warnings.len(), MAX_WARNINGS_PER_SHEET + 1);
        assert_eq!(
            parsed
                .warnings
                .iter()
                .filter(|w| w.as_str() == WARNINGS_SUPPRESSED_MESSAGE)
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression warning should be the last warning; warnings={:?}",
            parsed.warnings
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
                .filter(|w| w.as_str() == WARNINGS_SUPPRESSED_MESSAGE)
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            props.warnings
        );
        assert_eq!(
            props.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression warning should be the last warning; warnings={:?}",
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
                .filter(|w| w.as_str() == WARNINGS_SUPPRESSED_MESSAGE)
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            props.warnings
        );
        assert_eq!(
            props.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression warning should be the last warning; warnings={:?}",
            props.warnings
        );
    }

    #[test]
    fn manual_page_break_cap_warning_is_forced_when_warnings_full() {
        let mut warnings = Vec::new();
        // Fill the warning buffer and include the suppression marker to simulate prior noisy
        // best-effort parsing.
        for idx in 0..MAX_WARNINGS_PER_SHEET {
            warnings.push(format!("warning {idx}"));
        }
        warnings.push(WARNINGS_SUPPRESSED_MESSAGE.to_string());

        let mut breaks = ManualPageBreaks::default();

        // HorizontalPageBreaks payload:
        // - cbrk (u16)
        // - HorzBrk[cbrk] (6 bytes each): row (u16), colStart (u16), colEnd (u16)
        //
        // Declare an absurd cbrk but only provide a single entry so the parser triggers the cap
        // warning.
        let mut data = Vec::new();
        data.extend_from_slice(&u16::MAX.to_le_bytes()); // cbrk
        data.extend_from_slice(&2u16.to_le_bytes()); // row
        data.extend_from_slice(&0u16.to_le_bytes()); // colStart
        data.extend_from_slice(&0u16.to_le_bytes()); // colEnd

        parse_horizontal_page_breaks_record(&data, 123, &mut breaks, &mut warnings);

        assert!(
            warnings
                .iter()
                .any(|w| w.contains("HorizontalPageBreaks") && w.contains("cbrk=")),
            "expected forced cap warning, got {warnings:?}"
        );
        assert_eq!(
            warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should be preserved; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_sort_record_split_across_continue_records() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // A well-formed SORT record split across CONTINUE. This validates that worksheet parsers
        // use `LogicalBiffRecordIter` for continuable record ids (SORT can be continued).
        let mut sort_payload = vec![0u8; 24];
        // Sorted range: rows 0..=10, cols 0..=0.
        sort_payload[0..2].copy_from_slice(&0u16.to_le_bytes()); // rwFirst
        sort_payload[2..4].copy_from_slice(&10u16.to_le_bytes()); // rwLast
        sort_payload[4..6].copy_from_slice(&0u16.to_le_bytes()); // colFirst
        sort_payload[6..8].copy_from_slice(&0u16.to_le_bytes()); // colLast
        sort_payload[8..10].copy_from_slice(&0u16.to_le_bytes()); // grbit
        sort_payload[10..12].copy_from_slice(&1u16.to_le_bytes()); // cKey
        sort_payload[12..14].copy_from_slice(&0u16.to_le_bytes()); // key col 1
        sort_payload[14..16].copy_from_slice(&0xFFFFu16.to_le_bytes()); // key col 2 (unused)
        sort_payload[16..18].copy_from_slice(&0xFFFFu16.to_le_bytes()); // key col 3 (unused)
                                                                        // orders default to 0 (ascending).

        let split_at = 10usize;
        stream.extend_from_slice(&record(RECORD_SORT, &sort_payload[..split_at]));
        stream.extend_from_slice(&record(records::RECORD_CONTINUE, &sort_payload[split_at..]));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert_eq!(
            props.sort_state,
            Some(SortState {
                conditions: vec![SortCondition {
                    range: Range::new(CellRef::new(0, 0), CellRef::new(10, 0)),
                    descending: false,
                }]
            })
        );
    }

    #[test]
    fn parses_autofilter12_record_with_unicode_string_split_across_continuefrt12() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // AutoFilter12 future record storing one value ("ABCDE") for column 0.
        //
        // We split the UTF-16 character bytes across a ContinueFrt12 boundary. Excel inserts a
        // 1-byte continuation option flags prefix (fHighByte) at the start of the continued
        // fragment; without fragment-aware parsing that byte corrupts the UTF-16 stream.
        let mut autofilter12_payload = Vec::new();
        autofilter12_payload.extend_from_slice(&RECORD_AUTOFILTER12.to_le_bytes()); // FrtHeader.rt
        autofilter12_payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        autofilter12_payload.extend_from_slice(&0u32.to_le_bytes()); // reserved

        autofilter12_payload.extend_from_slice(&0u16.to_le_bytes()); // colId
        autofilter12_payload.extend_from_slice(&1u16.to_le_bytes()); // cVals

        // XLUnicodeString header.
        autofilter12_payload.extend_from_slice(&5u16.to_le_bytes()); // cch
        autofilter12_payload.push(STR_FLAG_HIGH_BYTE); // unicode

        // First two characters (A, B) stored in the AutoFilter12 record payload.
        autofilter12_payload.extend_from_slice(&(b'A' as u16).to_le_bytes());
        autofilter12_payload.extend_from_slice(&(b'B' as u16).to_le_bytes());

        // ContinueFrt12 payload begins with its own FrtHeader, then the continuation flags byte,
        // then the remaining characters (CDE).
        let mut cont_payload = Vec::new();
        cont_payload.extend_from_slice(&RECORD_CONTINUEFRT12.to_le_bytes()); // FrtHeader.rt
        cont_payload.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        cont_payload.extend_from_slice(&0u32.to_le_bytes()); // reserved
        cont_payload.push(STR_FLAG_HIGH_BYTE); // continuation option flags (unicode)
        cont_payload.extend_from_slice(&(b'C' as u16).to_le_bytes());
        cont_payload.extend_from_slice(&(b'D' as u16).to_le_bytes());
        cont_payload.extend_from_slice(&(b'E' as u16).to_le_bytes());

        stream.extend_from_slice(&record(RECORD_AUTOFILTER12, &autofilter12_payload));
        stream.extend_from_slice(&record(RECORD_CONTINUEFRT12, &cont_payload));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert!(
            props.warnings.is_empty(),
            "expected no warnings, got {:?}",
            props.warnings
        );

        assert_eq!(
            props.auto_filter_columns,
            vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Equals(FilterValue::Text(
                    "ABCDE".to_string()
                ))],
                values: vec!["ABCDE".to_string()],
                raw_xml: Vec::new(),
            }]
        );
    }

    #[test]
    fn sort_record_out_of_bounds_is_skipped_with_warning() {
        let mut stream = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // SORT record with an out-of-bounds col range (>= EXCEL_MAX_COLS).
        let mut sort_payload = vec![0u8; 24];
        sort_payload[0..2].copy_from_slice(&0u16.to_le_bytes()); // rwFirst
        sort_payload[2..4].copy_from_slice(&10u16.to_le_bytes()); // rwLast
        sort_payload[4..6].copy_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colFirst (OOB)
        sort_payload[6..8].copy_from_slice(&(EXCEL_MAX_COLS as u16).to_le_bytes()); // colLast (OOB)
        sort_payload[10..12].copy_from_slice(&1u16.to_le_bytes()); // cKey

        stream.extend_from_slice(&record(RECORD_SORT, &sort_payload));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let props = parse_biff_sheet_row_col_properties(&stream, 0, 1252).expect("parse");
        assert!(props.sort_state.is_none());
        assert!(
            props.warnings.iter().any(|w| {
                w.contains("failed to parse SORT record")
                    && w.contains("out of bounds")
                    && w.contains("offset")
            }),
            "expected out-of-bounds SORT warning, got {:?}",
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
                .filter(|w| w.as_str() == WARNINGS_SUPPRESSED_MESSAGE)
                .count(),
            1,
            "suppression warning should only be emitted once; warnings={:?}",
            links.warnings
        );
        assert_eq!(
            links.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression warning should be the last warning; warnings={:?}",
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

        let parsed = parse_biff_sheet_merged_cells(&stream, 0).expect("parse");
        assert!(
            parsed.warnings.is_empty(),
            "expected no warnings, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.ranges,
            vec![
                Range::from_a1("A1:B1").unwrap(),
                Range::from_a1("C2:D3").unwrap(),
            ]
        );
    }

    #[test]
    fn caps_mergedcells_ranges_per_sheet() {
        let cap = MAX_MERGED_RANGES_PER_SHEET;
        assert!(cap >= 2, "test requires cap >= 2");

        // Build more valid Ref8 areas than the hard cap:
        // - First record includes `cap - 1` ranges.
        // - Second record includes 2 ranges: one to reach the cap, plus a unique one beyond it.
        let first_count = cap - 1;
        let first_count_u16 =
            u16::try_from(first_count).expect("test cap must fit in a MERGEDCELLS record");

        let mut merged1 = Vec::new();
        merged1.extend_from_slice(&first_count_u16.to_le_bytes());
        for i in 0..first_count {
            let row = u16::try_from(i).expect("row should fit in u16");
            merged1.extend_from_slice(&row.to_le_bytes()); // rwFirst
            merged1.extend_from_slice(&row.to_le_bytes()); // rwLast
            merged1.extend_from_slice(&0u16.to_le_bytes()); // colFirst
            merged1.extend_from_slice(&1u16.to_le_bytes()); // colLast
        }

        let cap_row = first_count_u16;
        let beyond_row = 9999u16;
        let beyond_range = Range::new(
            CellRef::new(beyond_row as u32, 5),
            CellRef::new(beyond_row as u32, 6),
        );

        let mut merged2 = Vec::new();
        merged2.extend_from_slice(&2u16.to_le_bytes()); // cAreas
                                                        // Range that reaches the cap.
        merged2.extend_from_slice(&cap_row.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&cap_row.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&0u16.to_le_bytes()); // colFirst
        merged2.extend_from_slice(&1u16.to_le_bytes()); // colLast
                                                        // Unique range beyond the cap.
        merged2.extend_from_slice(&beyond_row.to_le_bytes()); // rwFirst
        merged2.extend_from_slice(&beyond_row.to_le_bytes()); // rwLast
        merged2.extend_from_slice(&5u16.to_le_bytes()); // colFirst
        merged2.extend_from_slice(&6u16.to_le_bytes()); // colLast

        let stream = [
            record(RECORD_MERGEDCELLS, &merged1),
            record(RECORD_MERGEDCELLS, &merged2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_merged_cells(&stream, 0).expect("parse");
        assert_eq!(parsed.ranges.len(), cap);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many merged ranges")),
            "expected cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.ranges.last().copied(),
            Some(Range::new(
                CellRef::new(cap_row as u32, 0),
                CellRef::new(cap_row as u32, 1),
            ))
        );
        assert!(
            !parsed.ranges.contains(&beyond_range),
            "expected range beyond cap to be absent"
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
    fn cell_xf_scan_refuses_to_allocate_unbounded_maps() {
        // Use a small cap so the test runs quickly and without large allocations.
        let cap = 10usize;

        let mut parts = Vec::new();
        for row in 0..(cap as u16 + 1) {
            let mut data = Vec::new();
            data.extend_from_slice(&row.to_le_bytes()); // row
            data.extend_from_slice(&0u16.to_le_bytes()); // col
            data.extend_from_slice(&1u16.to_le_bytes()); // xf
            parts.push(record(RECORD_BLANK, &data));
        }
        parts.push(record(records::RECORD_EOF, &[]));
        let stream = parts.concat();

        let err = parse_biff_sheet_cell_xf_indices_filtered_with_cap(&stream, 0, None, cap)
            .expect_err("expected cap error");
        assert!(
            err.contains("too many cell XF entries"),
            "unexpected error: {err}"
        );
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
    fn sheet_view_state_metadata_warnings_are_capped() {
        let mut stream: Vec<u8> = Vec::new();
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 25) {
            stream.extend_from_slice(&record(RECORD_WINDOW2, &[0u8; 1]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "expected warnings to be capped at {} (+1 suppression), got {}",
            MAX_WARNINGS_PER_SHEET_METADATA,
            parsed.warnings.len()
        );
        assert_eq!(
            parsed
                .warnings
                .iter()
                .filter(|w| w.as_str() == SHEET_METADATA_WARNINGS_SUPPRESSED)
                .count(),
            1,
            "expected exactly one suppression warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "expected suppression warning to be the final warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_view_state_scan_stops_after_record_cap() {
        let cap = MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN;
        assert!(cap >= 10, "test requires cap >= 10");

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..(cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        // This record should be ignored because the scan stops at the cap.
        stream.extend_from_slice(&record(RECORD_WINDOW2, &WINDOW2_FLAG_DSP_GRID.to_le_bytes()));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        assert_eq!(parsed.show_grid_lines, None);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("view state")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_view_state_record_cap_warning_is_emitted_even_when_other_metadata_warnings_are_suppressed()
    {
        let record_cap = MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN;
        assert!(
            record_cap > MAX_WARNINGS_PER_SHEET_METADATA + 10,
            "test requires record cap to exceed warning cap"
        );

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the metadata warning buffer with malformed WINDOW2 records (payload too short).
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 10) {
            stream.extend_from_slice(&record(RECORD_WINDOW2, &[0u8; 1]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("view state")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_view_state_selection_cap_warning_is_emitted_even_when_other_metadata_warnings_are_suppressed(
    ) {
        let selection_cap = MAX_SELECTION_RANGES_PER_RECORD;
        assert!(selection_cap >= 2, "test requires selection cap >= 2");

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the metadata warning buffer with malformed WINDOW2 records (payload too short).
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 10) {
            stream.extend_from_slice(&record(RECORD_WINDOW2, &[0u8; 1]));
        }

        // Emit a SELECTION record whose cref exceeds the selection-range cap.
        let pane: u8 = 0;
        let active_row: u16 = 0;
        let active_col: u16 = 0;
        let declared = selection_cap + 10;
        let declared_u16 = u16::try_from(declared).expect("cref should fit in u16");

        let mut selection_payload = Vec::new();
        selection_payload.push(pane);
        selection_payload.extend_from_slice(&active_row.to_le_bytes());
        selection_payload.extend_from_slice(&active_col.to_le_bytes());
        selection_payload.extend_from_slice(&0u16.to_le_bytes()); // irefActive (ignored)
        selection_payload.extend_from_slice(&declared_u16.to_le_bytes());
        for idx in 0..declared {
            let row = u16::try_from(idx).expect("row should fit in u16");
            selection_payload.extend_from_slice(&row.to_le_bytes()); // rwFirst
            selection_payload.extend_from_slice(&row.to_le_bytes()); // rwLast
            selection_payload.push(0); // colFirst
            selection_payload.push(0); // colLast
        }
        stream.extend_from_slice(&record(RECORD_SELECTION, &selection_payload));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("SELECTION record") && w.contains("cap=")),
            "expected forced selection-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_view_state_selection_record_cap_warning_is_emitted_even_when_other_metadata_warnings_are_suppressed(
    ) {
        let selection_record_cap = MAX_SELECTION_RECORDS_PER_SHEET_VIEW_STATE;
        assert!(selection_record_cap >= 2, "test requires selection record cap >= 2");

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the metadata warning buffer with malformed WINDOW2 records (payload too short).
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 10) {
            stream.extend_from_slice(&record(RECORD_WINDOW2, &[0u8; 1]));
        }

        // Emit more distinct pane selections than the retention cap.
        let active_row: u16 = 0;
        let active_col: u16 = 0;
        for pane in 0u8..u8::try_from(selection_record_cap + 5).unwrap() {
            let mut payload = Vec::new();
            payload.push(pane);
            payload.extend_from_slice(&active_row.to_le_bytes());
            payload.extend_from_slice(&active_col.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes()); // irefActive
            payload.extend_from_slice(&0u16.to_le_bytes()); // cref=0
            stream.extend_from_slice(&record(RECORD_SELECTION, &payload));
        }

        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many SELECTION records") && w.contains("cap=")),
            "expected forced selection-record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_view_state_selection_ranges_are_capped() {
        let cap = MAX_SELECTION_RANGES_PER_RECORD;
        assert!(cap >= 2, "test requires cap >= 2");

        // Build a SELECTION record using the PnnU8NoPadRefU layout (pane id u8, RefU ranges).
        let pane: u8 = 0;
        let active_row: u16 = 0;
        let active_col: u16 = 0;

        let declared = cap + 10;
        let declared_u16 = u16::try_from(declared).expect("test selection cref should fit in u16");

        let mut selection_payload = Vec::new();
        selection_payload.push(pane);
        selection_payload.extend_from_slice(&active_row.to_le_bytes());
        selection_payload.extend_from_slice(&active_col.to_le_bytes());
        selection_payload.extend_from_slice(&0u16.to_le_bytes()); // irefActive (ignored)
        selection_payload.extend_from_slice(&declared_u16.to_le_bytes());

        for idx in 0..declared {
            let row = u16::try_from(idx).expect("row should fit in u16");
            selection_payload.extend_from_slice(&row.to_le_bytes()); // rwFirst
            selection_payload.extend_from_slice(&row.to_le_bytes()); // rwLast
            selection_payload.push(0); // colFirst
            selection_payload.push(0); // colLast
        }

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SELECTION, &selection_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_sheet_view_state(&stream, 0).expect("parse");
        let selection = parsed.selection.expect("selection missing");

        assert_eq!(selection.active_cell, CellRef::new(0, 0));
        assert_eq!(selection.ranges.len(), cap);
        assert_eq!(
            selection.ranges.first().copied(),
            Some(Range::new(CellRef::new(0, 0), CellRef::new(0, 0)))
        );
        assert_eq!(
            selection.ranges.last().copied(),
            Some(Range::new(
                CellRef::new((cap - 1) as u32, 0),
                CellRef::new((cap - 1) as u32, 0)
            ))
        );

        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("SELECTION record") && w.contains("cap=")),
            "expected selection cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_scan_stops_after_record_cap() {
        let cap = MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN;
        assert!(cap >= 10, "test requires cap >= 10");

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..(cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        // This record should be ignored because the scan stops at the cap.
        stream.extend_from_slice(&record(RECORD_PROTECT, &1u16.to_le_bytes()));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(parsed.protection.enabled, false);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("sheet protection")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_record_cap_warning_is_emitted_even_when_other_metadata_warnings_are_suppressed()
    {
        let record_cap = MAX_RECORDS_SCANNED_PER_SHEET_METADATA_SCAN;
        assert!(
            record_cap > MAX_WARNINGS_PER_SHEET_METADATA + 10,
            "test requires record cap to exceed warning cap"
        );

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill warnings with truncated protection records.
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 10) {
            stream.extend_from_slice(&record(RECORD_PROTECT, &[1]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("sheet protection")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_protection_metadata_warnings_are_capped() {
        let mut stream: Vec<u8> = Vec::new();
        for _ in 0..(MAX_WARNINGS_PER_SHEET_METADATA + 25) {
            stream.extend_from_slice(&record(RECORD_PROTECT, &[0u8; 1]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_protection(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET_METADATA + 1,
            "expected warnings to be capped at {} (+1 suppression), got {}",
            MAX_WARNINGS_PER_SHEET_METADATA,
            parsed.warnings.len()
        );
        assert_eq!(
            parsed
                .warnings
                .iter()
                .filter(|w| w.as_str() == SHEET_METADATA_WARNINGS_SUPPRESSED)
                .count(),
            1,
            "expected exactly one suppression warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(SHEET_METADATA_WARNINGS_SUPPRESSED),
            "expected suppression warning to be the final warning, got {:?}",
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

    #[test]
    fn manual_page_break_scan_stops_after_record_cap() {
        let cap = MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN;
        assert!(cap >= 10, "test requires cap >= 10");

        let mut horizontal = Vec::new();
        horizontal.extend_from_slice(&1u16.to_le_bytes()); // cbrk
        horizontal.extend_from_slice(&2u16.to_le_bytes()); // row
        horizontal.extend_from_slice(&0u16.to_le_bytes()); // colStart
        horizontal.extend_from_slice(&0u16.to_le_bytes()); // colEnd

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..(cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        // This record should be ignored because the scan stops at the record cap.
        stream.extend_from_slice(&record(RECORD_HORIZONTALPAGEBREAKS, &horizontal));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_manual_page_breaks(&stream, 0).expect("parse");
        assert!(parsed.manual_page_breaks.row_breaks_after.is_empty());
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("manual page breaks")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn manual_page_break_record_cap_warning_is_forced_when_warnings_full() {
        let record_cap = MAX_RECORDS_SCANNED_PER_SHEET_PAGE_BREAK_SCAN;
        assert!(
            record_cap > MAX_WARNINGS_PER_SHEET + 10,
            "test requires record cap to exceed warning cap"
        );

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the warning buffer with truncated HorizontalPageBreaks records.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 10) {
            stream.extend_from_slice(&record(RECORD_HORIZONTALPAGEBREAKS, &[]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_manual_page_breaks(&stream, 0).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("manual page breaks")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_labelsst_scan_stops_after_record_cap() {
        let cap = MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN;
        assert!(cap >= 10, "test requires cap >= 10");

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // row
        payload.extend_from_slice(&0u16.to_le_bytes()); // col
        payload.extend_from_slice(&0u16.to_le_bytes()); // xf
        payload.extend_from_slice(&0u32.to_le_bytes()); // isst

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..(cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        // This record should be ignored because the scan stops at the record cap.
        stream.extend_from_slice(&record(RECORD_LABELSST, &payload));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let sst_phonetics = vec![Some("phonetic".to_string())];
        let parsed =
            parse_biff_sheet_labelsst_indices(&stream, 0, Some(sst_phonetics.as_slice()))
                .expect("parse");
        assert!(parsed.indices.is_empty());
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("LABELSST")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_labelsst_entries_are_capped() {
        let cap = MAX_LABELSST_ENTRIES_PER_SHEET;
        assert!(cap >= 2, "test requires cap >= 2");

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        for row in 0u16..u16::try_from(cap + 10).unwrap() {
            let mut payload = Vec::new();
            payload.extend_from_slice(&row.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes()); // col
            payload.extend_from_slice(&0u16.to_le_bytes()); // xf
            payload.extend_from_slice(&0u32.to_le_bytes()); // isst
            stream.extend_from_slice(&record(RECORD_LABELSST, &payload));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let sst_phonetics = vec![Some("phonetic".to_string())];
        let parsed =
            parse_biff_sheet_labelsst_indices(&stream, 0, Some(sst_phonetics.as_slice()))
                .expect("parse");
        assert_eq!(parsed.indices.len(), cap);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many LABELSST indices") && w.contains("cap=")),
            "expected entry-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_labelsst_record_cap_warning_is_emitted_even_when_warnings_are_suppressed() {
        let record_cap = MAX_RECORDS_SCANNED_PER_SHEET_LABELSST_SCAN;
        assert!(
            record_cap > MAX_WARNINGS_PER_SHEET + 10,
            "test requires record cap to exceed warning cap"
        );

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill warnings with malformed LABELSST records.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 10) {
            stream.extend_from_slice(&record(RECORD_LABELSST, &[0u8; 1]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_labelsst_indices(&stream, 0, None).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("LABELSST")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should remain last; warnings={:?}",
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
            let utf16le: Vec<u8> = location
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<u8>>();
            data.extend_from_slice(&((utf16le.len() / 2) as u32).to_le_bytes());
            data.extend_from_slice(&utf16le);

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

    #[test]
    fn sheet_hyperlink_cap_warning_is_emitted_even_when_other_warnings_are_suppressed() {
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
            let utf16le: Vec<u8> = location
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<u8>>();
            data.extend_from_slice(&((utf16le.len() / 2) as u32).to_le_bytes());
            data.extend_from_slice(&utf16le);

            data
        }

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // First, generate enough malformed HLINK records to fill and suppress the warnings buffer.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 100) {
            stream.extend_from_slice(&record(RECORD_HLINK, &[]));
        }

        // Then emit more HLINK records than the hyperlink cap. The truncation warning should still
        // be present even though the warning buffer is already full.
        for _ in 0..(MAX_HYPERLINKS_PER_SHEET + 1) {
            let payload = internal_hlink_payload("Sheet1!A1");
            stream.extend_from_slice(&record(RECORD_HLINK, &payload));
        }

        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");
        assert_eq!(parsed.hyperlinks.len(), MAX_HYPERLINKS_PER_SHEET);
        assert_eq!(parsed.warnings.len(), MAX_WARNINGS_PER_SHEET + 1);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w == "too many hyperlinks; additional HLINK records skipped"),
            "expected hyperlink truncation warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed
                .warnings
                .iter()
                .filter(|w| w.as_str() == WARNINGS_SUPPRESSED_MESSAGE)
                .count(),
            1,
            "expected exactly one suppression warning; warnings={:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(|w| w.as_str()),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "expected suppression message to remain last; warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_hyperlink_scan_breaks_before_decoding_additional_hlink_records() {
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
            let utf16le: Vec<u8> = location
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<u8>>();
            data.extend_from_slice(&((utf16le.len() / 2) as u32).to_le_bytes());
            data.extend_from_slice(&utf16le);

            data
        }

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill exactly up to the cap with valid hyperlinks.
        for _ in 0..MAX_HYPERLINKS_PER_SHEET {
            let payload = internal_hlink_payload("Sheet1!A1");
            stream.extend_from_slice(&record(RECORD_HLINK, &payload));
        }

        // Next HLINK record is malformed; if we attempted to decode it we would emit a warning.
        // The parser should stop *before* decoding any additional HLINK records after reaching the
        // cap, so we should only see the truncation warning.
        stream.extend_from_slice(&record(RECORD_HLINK, &[]));

        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");
        assert_eq!(parsed.hyperlinks.len(), MAX_HYPERLINKS_PER_SHEET);
        assert_eq!(
            parsed.warnings,
            vec!["too many hyperlinks; additional HLINK records skipped".to_string()]
        );
    }

    #[test]
    fn sheet_hyperlink_scan_stops_after_record_cap() {
        let cap = MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN;
        assert!(cap >= 10, "test requires cap >= 10");

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
            let utf16le: Vec<u8> = location
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<u8>>();
            data.extend_from_slice(&((utf16le.len() / 2) as u32).to_le_bytes());
            data.extend_from_slice(&utf16le);

            data
        }

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));
        for _ in 0..(cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        // This record should be ignored because the scan stops at the record cap.
        stream.extend_from_slice(&record(RECORD_HLINK, &internal_hlink_payload("Sheet1!A1")));
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");
        assert_eq!(parsed.hyperlinks.len(), 0);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("hyperlinks")),
            "expected record-cap warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn sheet_hyperlink_record_cap_warning_is_emitted_even_when_other_warnings_are_suppressed() {
        let record_cap = MAX_RECORDS_SCANNED_PER_SHEET_HYPERLINK_SCAN;
        assert!(
            record_cap > MAX_WARNINGS_PER_SHEET + 10,
            "test requires record cap to exceed warning cap"
        );

        let mut stream: Vec<u8> = Vec::new();
        stream.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &[0u8; 16]));

        // Fill the warning buffer with malformed HLINK records (payload too short).
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 10) {
            stream.extend_from_slice(&record(RECORD_HLINK, &[]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&record(records::RECORD_EOF, &[]));

        let parsed = parse_biff_sheet_hyperlinks(&stream, 0, 1252).expect("parse");
        assert_eq!(
            parsed.warnings.len(),
            MAX_WARNINGS_PER_SHEET + 1,
            "warnings should remain capped; warnings={:?}",
            parsed.warnings
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("hyperlinks")),
            "expected forced record-cap warning, got {:?}",
            parsed.warnings
        );
        assert_eq!(
            parsed.warnings.last().map(String::as_str),
            Some(WARNINGS_SUPPRESSED_MESSAGE),
            "suppression marker should remain last; warnings={:?}",
            parsed.warnings
        );
    }
} 
