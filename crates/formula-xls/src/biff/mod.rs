//! Minimal BIFF record parsing helpers used by the legacy `.xls` importer.
//!
//! This module is intentionally best-effort: BIFF is large and this importer only
//! needs a handful of workbook-global and worksheet records. The implementation is
//! split into submodules so future BIFF parsing work can share common primitives
//! (record iteration, CONTINUE handling, and string decoding).
//!
//! In particular, this includes best-effort parsing for legacy cell notes/comments
//! (`NOTE/OBJ/TXO`) used by the `.xls` importer.

use std::io::{Read, Seek};
use std::path::Path;

pub(crate) mod autofilter;
pub(crate) mod autofilter_criteria;
mod comments;
pub(crate) mod defined_names;
pub(crate) mod encryption;
pub(crate) mod externsheet;
pub(crate) mod formulas;
pub(crate) mod globals;
pub(crate) mod print_settings;
pub(crate) mod workbook_context;
pub(crate) mod rgce;
pub(crate) mod records;
pub(crate) mod sheet;
pub(crate) mod shared_formulas;
pub(crate) mod sort;
pub(crate) mod sst;
pub(crate) mod strings;
pub(crate) mod supbook;
pub(crate) mod worksheet_formulas;

pub(crate) use autofilter::{parse_biff_filter_database_ranges, ParsedFilterDatabaseRanges};
pub(crate) use autofilter_criteria::parse_biff_sheet_autofilter_criteria;
pub(crate) use comments::parse_biff_sheet_notes;
pub(crate) use defined_names::parse_biff_defined_names;
pub(crate) use globals::{
    parse_biff_bound_sheets, parse_biff_codepage, parse_biff_workbook_globals, BoundSheetInfo,
};
pub(crate) use print_settings::parse_biff_sheet_print_settings;
pub(crate) use sheet::{
    parse_biff8_sheet_formulas,
    parse_biff8_sheet_table_formulas,
    parse_biff_sheet_cell_xf_indices_filtered,
    parse_biff_sheet_hyperlinks,
    parse_biff_sheet_labelsst_indices,
    parse_biff_sheet_merged_cells,
    parse_biff_sheet_protection,
    parse_biff_sheet_row_col_properties,
    parse_biff_sheet_view_state,
    SheetRowColProperties,
};
pub(crate) use sort::parse_biff_sheet_sort_state;
pub(crate) use shared_formulas::parse_biff_sheet_shared_formulas;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BiffVersion {
    Biff5,
    Biff8,
}

// BIFF version numbers stored in the BOF record payload.
// See [MS-XLS] 2.4.21 (BOF).
const BOF_VERSION_BIFF5: u16 = 0x0500;
const BOF_VERSION_BIFF8: u16 = 0x0600;
// BOF "substream type" value used by the `calamine` heuristic when the BIFF version is 0.
// 0x1000 corresponds to a worksheet substream.
const BOF_DT_WORKSHEET: u16 = 0x1000;

/// Read the workbook stream bytes from a compound file.
pub(crate) fn read_workbook_stream_from_xls(path: &Path) -> Result<Vec<u8>, String> {
    let mut comp = cfb::open(path).map_err(|err| err.to_string())?;
    let mut stream = open_xls_workbook_stream(&mut comp)?;

    let mut workbook_stream = Vec::new();
    stream
        .read_to_end(&mut workbook_stream)
        .map_err(|err| err.to_string())?;
    Ok(workbook_stream)
}

pub(crate) fn open_xls_workbook_stream<R: Read + Seek>(
    comp: &mut cfb::CompoundFile<R>,
) -> Result<cfb::Stream<R>, String> {
    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(stream) = comp.open_stream(candidate) {
            return Ok(stream);
        }
    }
    Err("missing workbook stream (expected `Workbook` or `Book`)".to_string())
}

pub(crate) fn detect_biff_version(workbook_stream: &[u8]) -> BiffVersion {
    let Some((record_id, data)) = records::read_biff_record(workbook_stream, 0) else {
        return BiffVersion::Biff8;
    };

    // BOF record type. Use BIFF8 heuristics compatible with calamine.
    if !records::is_bof_record(record_id) {
        return BiffVersion::Biff8;
    }

    let Some(biff_version) = data.get(0..2).map(|v| u16::from_le_bytes([v[0], v[1]])) else {
        return BiffVersion::Biff8;
    };

    let dt = data
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .unwrap_or(0);

    match biff_version {
        BOF_VERSION_BIFF5 => BiffVersion::Biff5,
        BOF_VERSION_BIFF8 => BiffVersion::Biff8,
        0 => {
            if dt == BOF_DT_WORKSHEET {
                BiffVersion::Biff5
            } else {
                BiffVersion::Biff8
            }
        }
        _ => BiffVersion::Biff8,
    }
}

#[cfg(test)]
mod fuzz_tests;

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    #[test]
    fn detects_biff8_from_bof_version() {
        let payload = [BOF_VERSION_BIFF8.to_le_bytes(), 0u16.to_le_bytes()].concat();
        let stream = record(records::RECORD_BOF_BIFF8, &payload);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }

    #[test]
    fn detects_biff5_from_bof_version() {
        let payload = [BOF_VERSION_BIFF5.to_le_bytes(), 0u16.to_le_bytes()].concat();
        let stream = record(records::RECORD_BOF_BIFF5, &payload);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff5);
    }

    #[test]
    fn detects_biff5_from_dt_heuristic_when_version_is_zero() {
        // BIFF version=0, dt=worksheet => BIFF5 heuristic.
        let payload = [0u16.to_le_bytes(), BOF_DT_WORKSHEET.to_le_bytes()].concat();
        let stream = record(records::RECORD_BOF_BIFF5, &payload);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff5);
    }

    #[test]
    fn defaults_to_biff8_when_version_is_zero_and_dt_is_not_worksheet() {
        let payload = [0u16.to_le_bytes(), 0u16.to_le_bytes()].concat();
        let stream = record(records::RECORD_BOF_BIFF5, &payload);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }

    #[test]
    fn defaults_to_biff8_for_missing_bof() {
        let stream = record(0x0001, &[0x00]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }
}
