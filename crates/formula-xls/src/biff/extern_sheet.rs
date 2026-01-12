//! BIFF8 `EXTERNSHEET` (0x0017) record parsing.
//!
//! Defined-name (`NAME`) formula token streams frequently contain `PtgRef3d` / `PtgArea3d` tokens
//! that reference an `ixti` entry in the workbook-global `EXTERNSHEET` table. This module provides
//! a minimal, best-effort parser for that table.

#![allow(dead_code)]

use super::records;

/// BIFF8 `EXTERNSHEET` record id.
///
/// See [MS-XLS] 2.4.102 (EXTERNSHEET).
const RECORD_EXTERNSHEET: u16 = 0x0017;

/// A best-effort decoded `XTI` entry from the `EXTERNSHEET` table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ExternSheetRef {
    /// Sheet index range within the *current* workbook (`iSupBook == 0`).
    ///
    /// `itab_first` and `itab_last` are 0-based BIFF sheet indices (BoundSheet order).
    Internal { itab_first: i16, itab_last: i16 },
    /// Reference to an external workbook / add-in (`iSupBook != 0`).
    ///
    /// This is not yet supported; callers should treat it as `#REF!`.
    External,
}

impl ExternSheetRef {
    pub(crate) fn tab_range(self) -> Option<(i16, i16)> {
        match self {
            ExternSheetRef::Internal {
                itab_first,
                itab_last,
            } => Some((itab_first, itab_last)),
            ExternSheetRef::External => None,
        }
    }
}

/// Best-effort parse result for the workbook-global `EXTERNSHEET` table.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct ExternSheetTable {
    /// Entries indexed by `ixti` (0-based).
    pub(crate) entries: Vec<ExternSheetRef>,
    /// Any non-fatal parse warnings.
    pub(crate) warnings: Vec<String>,
}

/// Scan the workbook-global BIFF substream for an `EXTERNSHEET` record and parse its XTI table.
///
/// Best-effort semantics:
/// - Stops at the workbook-global `EOF` record, or the next `BOF` record (start of the next
///   substream).
/// - If the record is truncated or malformed, emits a warning and returns what was parsed.
pub(crate) fn parse_biff8_externsheet_table(workbook_stream: &[u8]) -> ExternSheetTable {
    let mut out = ExternSheetTable::default();

    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);
    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                out.warnings.push(format!(
                    "malformed BIFF record while scanning for EXTERNSHEET: {err}"
                ));
                break;
            }
        };

        // Stop scanning at the start of the next substream (worksheet BOF), even if the workbook
        // globals are missing the expected EOF record.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_EXTERNSHEET => {
                parse_externsheet_record(&mut out, record.data.as_ref(), record.offset);
                break;
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    out
}

fn allows_continuation(record_id: u16) -> bool {
    // EXTERNSHEET can be large and may be split across one or more `CONTINUE` records.
    record_id == RECORD_EXTERNSHEET
}

fn parse_externsheet_record(out: &mut ExternSheetTable, data: &[u8], offset: usize) {
    if data.len() < 2 {
        out.warnings.push(format!(
            "truncated EXTERNSHEET record at offset {offset}: missing cxti"
        ));
        return;
    }

    let cxti = u16::from_le_bytes([data[0], data[1]]) as usize;
    let mut cursor = 2usize;

    out.entries.reserve(cxti);

    for parsed in 0..cxti {
        if data.len() < cursor + 6 {
            out.warnings.push(format!(
                "truncated EXTERNSHEET record at offset {offset}: expected {cxti} XTI entries, got {parsed}"
            ));
            break;
        }

        let i_sup_book = u16::from_le_bytes([data[cursor], data[cursor + 1]]);
        let itab_first = i16::from_le_bytes([data[cursor + 2], data[cursor + 3]]);
        let itab_last = i16::from_le_bytes([data[cursor + 4], data[cursor + 5]]);
        cursor += 6;

        if i_sup_book == 0 {
            out.entries.push(ExternSheetRef::Internal {
                itab_first,
                itab_last,
            });
        } else {
            // External workbook/add-in reference: leave as placeholder for `#REF!`.
            out.entries.push(ExternSheetRef::External);
        }
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
    fn parses_externsheet_table_internal_entries() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&2u16.to_le_bytes()); // cxti

        // ixti=0: iSupBook=0, itabFirst=0, itabLast=0
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());

        // ixti=1: iSupBook=0, itabFirst=1, itabLast=3
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&1i16.to_le_bytes());
        payload.extend_from_slice(&3i16.to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_externsheet_table(&stream);
        assert_eq!(
            parsed.entries,
            vec![
                ExternSheetRef::Internal {
                    itab_first: 0,
                    itab_last: 0
                },
                ExternSheetRef::Internal {
                    itab_first: 1,
                    itab_last: 3
                }
            ]
        );
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn coalesces_continues_for_externsheet_record() {
        let mut full = Vec::new();
        full.extend_from_slice(&2u16.to_le_bytes()); // cxti

        // ixti=0: internal sheet 0.
        full.extend_from_slice(&0u16.to_le_bytes());
        full.extend_from_slice(&0i16.to_le_bytes());
        full.extend_from_slice(&0i16.to_le_bytes());

        // ixti=1: internal sheet range 1..2.
        full.extend_from_slice(&0u16.to_le_bytes());
        full.extend_from_slice(&1i16.to_le_bytes());
        full.extend_from_slice(&2i16.to_le_bytes());

        // Split after the first entry.
        let first_part = &full[..(2 + 6)];
        let second_part = &full[(2 + 6)..];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, first_part),
            record(records::RECORD_CONTINUE, second_part),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_externsheet_table(&stream);
        assert_eq!(parsed.entries.len(), 2);
        assert!(
            parsed.entries[0].tab_range() == Some((0, 0)),
            "entry0={:?}",
            parsed.entries[0]
        );
        assert!(
            parsed.entries[1].tab_range() == Some((1, 2)),
            "entry1={:?}",
            parsed.entries[1]
        );
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn external_supbook_entries_are_marked_external() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes()); // cxti
        payload.extend_from_slice(&2u16.to_le_bytes()); // iSupBook != 0 => external
        payload.extend_from_slice(&0i16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_externsheet_table(&stream);
        assert_eq!(parsed.entries, vec![ExternSheetRef::External]);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn warns_and_returns_partial_on_truncated_payload() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&3u16.to_le_bytes()); // cxti says 3 entries

        // Provide only 2 entries worth of data.
        for itab in [0i16, 1i16] {
            payload.extend_from_slice(&0u16.to_le_bytes()); // iSupBook
            payload.extend_from_slice(&itab.to_le_bytes());
            payload.extend_from_slice(&itab.to_le_bytes());
        }

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_externsheet_table(&stream);
        assert_eq!(parsed.entries.len(), 2);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("truncated EXTERNSHEET")),
            "expected truncated warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn scan_stops_at_next_bof_without_eof() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());
        payload.extend_from_slice(&0i16.to_le_bytes());

        // EXTERNSHEET lives after the next BOF; it should be ignored.
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
        ]
        .concat();

        let parsed = parse_biff8_externsheet_table(&stream);
        assert!(parsed.entries.is_empty(), "entries={:?}", parsed.entries);
    }
}

