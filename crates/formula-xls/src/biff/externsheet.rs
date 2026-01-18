//! BIFF `EXTERNSHEET` (0x0017) record parsing.
//!
//! BIFF8 defined-name (`NAME`) formula token streams frequently contain `PtgRef3d` / `PtgArea3d`
//! tokens that reference an `ixti` entry in the workbook-global `EXTERNSHEET` table. This module
//! provides a small, best-effort parser for that table.

#![allow(dead_code)]

use super::{records, BiffVersion};

/// BIFF8 `EXTERNSHEET` record id.
///
/// See [MS-XLS] 2.4.102 (EXTERNSHEET).
const RECORD_EXTERNSHEET: u16 = 0x0017;

/// Hard cap on the number of XTI entries we will parse from an `EXTERNSHEET` record.
///
/// Corrupt or adversarial files can construct extremely large EXTERNSHEET tables that are not
/// useful for formula rendering but can consume significant memory. This conservative cap keeps
/// parsing and allocation bounded while remaining far above typical real-world workbooks.
const MAX_XTI_ENTRIES: usize = 16_384;

/// An entry in the BIFF8 `EXTERNSHEET` table.
///
/// This corresponds to one `XTI` structure in [MS-XLS] 2.4.102.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ExternSheetEntry {
    /// Index of the referenced `SUPBOOK` record (`iSupBook`).
    ///
    /// `0` indicates an internal workbook reference.
    pub(crate) supbook: u16,
    /// First BIFF sheet index in the referenced sheet range (`itabFirst`).
    pub(crate) itab_first: i16,
    /// Last BIFF sheet index in the referenced sheet range (`itabLast`).
    pub(crate) itab_last: i16,
}

/// Best-effort parse result for the workbook-global `EXTERNSHEET` table.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct ExternSheetTable {
    /// Entries indexed by `ixti` (0-based).
    pub(crate) entries: Vec<ExternSheetEntry>,
    /// Any non-fatal parse warnings.
    pub(crate) warnings: Vec<String>,
}

/// Scan the workbook-global BIFF substream for an `EXTERNSHEET` record and parse its XTI table.
///
/// Best-effort semantics:
/// - Stops at the workbook-global `EOF` record, or the next `BOF` record (start of the next
///   substream).
/// - If the record is truncated or malformed, emits a warning and returns what was parsed.
pub(crate) fn parse_biff_externsheet(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> ExternSheetTable {
    // Signature matches other workbook-global parsers; `codepage` is currently unused.
    let _ = codepage;

    match biff {
        BiffVersion::Biff8 => parse_biff8_externsheet_table(workbook_stream),
        // BIFF5 `EXTERNSHEET` is not currently needed by the importer.
        BiffVersion::Biff5 => ExternSheetTable::default(),
    }
}

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

/// Parse the payload bytes of a BIFF8 `EXTERNSHEET` record into an [`ExternSheetTable`].
///
/// This is useful for callers that already have the logical record data (e.g. from
/// [`records::LogicalBiffRecordIter`]) and want to share the same entry decoding + best-effort
/// warning behavior as [`parse_biff8_externsheet_table`].
pub(crate) fn parse_biff8_externsheet_record_data(data: &[u8], offset: usize) -> ExternSheetTable {
    let mut out = ExternSheetTable::default();
    parse_externsheet_record(&mut out, data, offset);
    out
}

fn allows_continuation(record_id: u16) -> bool {
    // EXTERNSHEET can be large and may be split across one or more `CONTINUE` records.
    record_id == RECORD_EXTERNSHEET
}

fn parse_externsheet_record(out: &mut ExternSheetTable, data: &[u8], offset: usize) {
    // BIFF8 EXTERNSHEET layout:
    //   [cXTI: u16]
    //   cXTI * [iSupBook: u16, itabFirst: i16, itabLast: i16]
    if data.len() < 2 {
        out.warnings.push(format!(
            "truncated EXTERNSHEET record at offset {offset}: missing cxti"
        ));
        return;
    }

    let cxti = u16::from_le_bytes([data[0], data[1]]) as usize;
    let max_entries = (data.len().saturating_sub(2)) / 6;
    let mut cursor = 2usize;

    if cxti > max_entries {
        out.warnings.push(format!(
            "EXTERNSHEET cxti={cxti} exceeds available data; clamping to {max_entries}"
        ));
    }

    let mut to_parse = cxti.min(max_entries);
    if to_parse > MAX_XTI_ENTRIES {
        out.warnings.push(format!(
            "EXTERNSHEET has {to_parse} XTI entries; capping to {MAX_XTI_ENTRIES}"
        ));
        to_parse = MAX_XTI_ENTRIES;
    }

    let _ = out.entries.try_reserve_exact(to_parse);

    for _ in 0..to_parse {
        let entry_end = match cursor.checked_add(6) {
            Some(v) => v,
            None => {
                debug_assert!(
                    false,
                    "EXTERNSHEET cursor overflow (cursor={cursor}, len={})",
                    data.len()
                );
                out.warnings.push(format!(
                    "truncated EXTERNSHEET record at offset {offset}: XTI entry cursor overflow (cursor={cursor}, len={})",
                    data.len()
                ));
                break;
            }
        };
        let Some(chunk) = data.get(cursor..entry_end) else {
            debug_assert!(
                false,
                "EXTERNSHEET cursor out of bounds (cursor={cursor}, len={})",
                data.len()
            );
            out.warnings.push(format!(
                "truncated EXTERNSHEET record at offset {offset}: missing XTI entry bytes (cursor={cursor}, len={})",
                data.len()
            ));
            break;
        };
        let supbook = u16::from_le_bytes([chunk[0], chunk[1]]);
        let itab_first = i16::from_le_bytes([chunk[2], chunk[3]]);
        let itab_last = i16::from_le_bytes([chunk[4], chunk[5]]);
        cursor = entry_end;

        out.entries.push(ExternSheetEntry {
            supbook,
            itab_first,
            itab_last,
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn externsheet_payload(entries: &[(u16, u16, u16)]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        for &(supbook, itab_first, itab_last) in entries {
            payload.extend_from_slice(&supbook.to_le_bytes());
            payload.extend_from_slice(&(itab_first as i16).to_le_bytes());
            payload.extend_from_slice(&(itab_last as i16).to_le_bytes());
        }
        payload
    }

    #[test]
    fn parses_externsheet_entries_biff8() {
        let entries = [(0, 2, 3), (0, 5, 6)];
        let payload = externsheet_payload(&entries);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert_eq!(
            parsed.entries,
            vec![
                ExternSheetEntry {
                    supbook: 0,
                    itab_first: 2,
                    itab_last: 3,
                },
                ExternSheetEntry {
                    supbook: 0,
                    itab_first: 5,
                    itab_last: 6,
                }
            ]
        );
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn parses_externsheet_across_continue_fragment_boundaries() {
        let entries = [(0x0010, 0x0020, 0x0030), (0x0040, 0x0050, 0x0060)];
        let payload = externsheet_payload(&entries);

        // Split the payload so a u16 value spans the EXTERNSHEET/CONTINUE boundary.
        let split = 2 + 6 + 1; // cXTI + first entry + 1 byte of second supbook
        let first = &payload[..split];
        let second = &payload[split..];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, first),
            record(records::RECORD_CONTINUE, second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert_eq!(
            parsed.entries,
            vec![
                ExternSheetEntry {
                    supbook: 0x0010,
                    itab_first: 0x0020,
                    itab_last: 0x0030,
                },
                ExternSheetEntry {
                    supbook: 0x0040,
                    itab_first: 0x0050,
                    itab_last: 0x0060,
                }
            ]
        );
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn scan_stops_at_next_bof_without_eof() {
        let entries = [(1, 2, 3), (4, 5, 6)];
        let payload = externsheet_payload(&entries);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            // Start of first worksheet substream.
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert_eq!(parsed.entries.len(), 2);
        assert_eq!(parsed.entries[0].itab_first, 2);
        assert_eq!(parsed.entries[1].itab_last, 6);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn scan_stops_on_malformed_record_and_returns_partial() {
        // Truncated record: declares 4 bytes but only provides 2. Must be at end of stream.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x1234u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[0xAA, 0xBB]);

        let stream = [record(records::RECORD_BOF_BIFF8, &[0u8; 16]), truncated].concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert!(
            parsed.entries.is_empty(),
            "expected empty table on malformed record, got {:?}",
            parsed.entries
        );
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("malformed BIFF record")),
            "expected malformed warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn scan_ignores_externsheet_after_next_bof_without_eof() {
        let entries = [(0, 0, 0)];
        let payload = externsheet_payload(&entries);

        // EXTERNSHEET lives after the next BOF; it should be ignored.
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert!(parsed.entries.is_empty(), "entries={:?}", parsed.entries);
    }

    #[test]
    fn warns_and_returns_partial_on_truncated_payload() {
        let entries = [(0, 0, 0), (0, 1, 1), (0, 2, 2)];
        let payload = externsheet_payload(&entries);

        // Truncate so we only have 2 entries worth of data.
        let truncated = &payload[..(2 + 6 * 2)];

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, truncated),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert_eq!(parsed.entries.len(), 2);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("clamping") && w.contains("EXTERNSHEET cxti=")),
            "expected clamping warning, got {:?}",
            parsed.warnings
        );
    }

    #[test]
    fn preserves_supbook_for_external_references() {
        let entries = [(2, 0, 0)];
        let payload = externsheet_payload(&entries);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252);
        assert_eq!(
            parsed.entries,
            vec![ExternSheetEntry {
                supbook: 2,
                itab_first: 0,
                itab_last: 0,
            }]
        );
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
    }

    #[test]
    fn clamps_absurd_cxti_without_allocating() {
        // Corrupt file: declares an absurd cxti but only provides 1 complete entry worth of bytes.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cXTI = 65535
                                                             // One XTI entry.
        payload.extend_from_slice(&0u16.to_le_bytes()); // iSupBook
        payload.extend_from_slice(&0i16.to_le_bytes()); // itabFirst
        payload.extend_from_slice(&0i16.to_le_bytes()); // itabLast

        let parsed = parse_biff8_externsheet_record_data(&payload, 0);
        assert_eq!(parsed.entries.len(), 1);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w.contains("clamping") && w.contains("65535")),
            "expected clamping warning, got {:?}",
            parsed.warnings
        );
        // Ensure we didn't reserve/allocate anywhere near 65535 entries.
        assert!(
            parsed.entries.capacity() < 1024,
            "unexpectedly large capacity={}",
            parsed.entries.capacity()
        );
    }
}
