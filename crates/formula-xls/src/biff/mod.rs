//! Minimal BIFF record parsing helpers used by the legacy `.xls` importer.
//!
//! This module is intentionally best-effort: BIFF is large and this importer only
//! needs a handful of workbook-global and worksheet records. The implementation is
//! split into submodules so future BIFF parsing work can share common primitives
//! (record iteration, CONTINUE handling, and string decoding).

use std::io::{Read, Seek};
use std::path::Path;

pub(crate) mod globals;
pub(crate) mod records;
pub(crate) mod sheet;
pub(crate) mod strings;

pub(crate) use globals::{parse_biff_bound_sheets, parse_biff_workbook_globals, BoundSheetInfo};
pub(crate) use sheet::{
    parse_biff_sheet_cell_xf_indices_filtered, parse_biff_sheet_row_col_properties,
    SheetRowColProperties,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BiffVersion {
    Biff5,
    Biff8,
}

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
        0x0500 => BiffVersion::Biff5,
        0x0600 => BiffVersion::Biff8,
        0 => {
            if dt == 0x1000 {
                BiffVersion::Biff5
            } else {
                BiffVersion::Biff8
            }
        }
        _ => BiffVersion::Biff8,
    }
}

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
        let stream = record(records::RECORD_BOF_BIFF8, &[0x00, 0x06, 0x00, 0x00]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }

    #[test]
    fn detects_biff5_from_bof_version() {
        let stream = record(records::RECORD_BOF_BIFF5, &[0x00, 0x05, 0x00, 0x00]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff5);
    }

    #[test]
    fn detects_biff5_from_dt_heuristic_when_version_is_zero() {
        // BIFF version=0, dt=0x1000 => BIFF5 heuristic.
        let stream = record(records::RECORD_BOF_BIFF5, &[0x00, 0x00, 0x00, 0x10]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff5);
    }

    #[test]
    fn defaults_to_biff8_when_version_is_zero_and_dt_is_not_worksheet() {
        let stream = record(records::RECORD_BOF_BIFF5, &[0x00, 0x00, 0x00, 0x00]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }

    #[test]
    fn defaults_to_biff8_for_missing_bof() {
        let stream = record(0x0001, &[0x00]);
        assert_eq!(detect_biff_version(&stream), BiffVersion::Biff8);
    }
}
