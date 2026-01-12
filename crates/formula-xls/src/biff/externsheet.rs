#![allow(dead_code)]

use super::{records, BiffVersion};

/// BIFF8 `EXTERNSHEET` record id.
///
/// See [MS-XLS] 2.4.102 (EXTERNSHEET).
const RECORD_EXTERNSHEET: u16 = 0x0017;

/// An entry in the BIFF8 `EXTERNSHEET` table.
///
/// This corresponds to one `XTI` structure in [MS-XLS] 2.4.102.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ExternSheetEntry {
    pub(crate) supbook: u16,
    pub(crate) itab_first: u16,
    pub(crate) itab_last: u16,
}

pub(crate) fn parse_biff_externsheet(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<Vec<ExternSheetEntry>, String> {
    // Signature matches other workbook-global parsers; `codepage` is currently unused.
    let _ = codepage;

    match biff {
        BiffVersion::Biff8 => parse_biff8_externsheet(workbook_stream),
        // BIFF5 `EXTERNSHEET` is not currently needed by the importer.
        BiffVersion::Biff5 => Ok(Vec::new()),
    }
}

fn parse_biff8_externsheet(workbook_stream: &[u8]) -> Result<Vec<ExternSheetEntry>, String> {
    let mut out = Vec::new();

    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);
    for record in iter {
        let record = match record {
            Ok(record) => record,
            // Best-effort: stop at the first malformed/truncated physical record and return what
            // we've parsed so far.
            Err(_) => break,
        };

        // The `EXTERNSHEET` record lives in the workbook-global substream. Stop if we see the start
        // of the next substream (worksheet BOF), even if the workbook-global EOF is missing.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_EXTERNSHEET => {
                out.extend(parse_biff8_externsheet_record(&record));
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn allows_continuation(record_id: u16) -> bool {
    record_id == RECORD_EXTERNSHEET
}

fn parse_biff8_externsheet_record(record: &records::LogicalBiffRecord<'_>) -> Vec<ExternSheetEntry> {
    // BIFF8 EXTERNSHEET layout:
    //   [cXTI: u16]
    //   cXTI * [iSupBook: u16, itabFirst: u16, itabLast: u16]
    //
    // The record may be split across one or more `CONTINUE` records; use the logical record
    // fragments so we can read cleanly across boundaries.
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let total_len: usize = fragments.iter().map(|f| f.len()).sum();

    let mut cursor = FragmentCursor::new(&fragments);

    let cxti = match cursor.read_u16_le() {
        Ok(v) => v as usize,
        Err(_) => return Vec::new(),
    };

    let available_entries = total_len.saturating_sub(2) / 6;
    let count = cxti.min(available_entries);

    let mut out = Vec::with_capacity(count);
    for _ in 0..count {
        let supbook = match cursor.read_u16_le() {
            Ok(v) => v,
            Err(_) => break,
        };
        let itab_first = match cursor.read_u16_le() {
            Ok(v) => v,
            Err(_) => break,
        };
        let itab_last = match cursor.read_u16_le() {
            Ok(v) => v,
            Err(_) => break,
        };
        out.push(ExternSheetEntry {
            supbook,
            itab_first,
            itab_last,
        });
    }

    out
}

struct FragmentCursor<'a> {
    fragments: &'a [&'a [u8]],
    frag_idx: usize,
    offset: usize,
}

impl<'a> FragmentCursor<'a> {
    fn new(fragments: &'a [&'a [u8]]) -> Self {
        Self {
            fragments,
            frag_idx: 0,
            offset: 0,
        }
    }

    fn advance_fragment(&mut self) -> Result<(), ()> {
        self.frag_idx = self.frag_idx.saturating_add(1);
        self.offset = 0;
        if self.frag_idx >= self.fragments.len() {
            return Err(());
        }
        Ok(())
    }

    fn read_u8(&mut self) -> Result<u8, ()> {
        loop {
            let frag = self.fragments.get(self.frag_idx).ok_or(())?;
            if self.offset < frag.len() {
                let b = frag[self.offset];
                self.offset += 1;
                return Ok(b);
            }
            self.advance_fragment()?;
        }
    }

    fn read_u16_le(&mut self) -> Result<u16, ()> {
        let lo = self.read_u8()?;
        let hi = self.read_u8()?;
        Ok(u16::from_le_bytes([lo, hi]))
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

    fn externsheet_payload(entries: &[(u16, u16, u16)]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        for &(supbook, itab_first, itab_last) in entries {
            payload.extend_from_slice(&supbook.to_le_bytes());
            payload.extend_from_slice(&itab_first.to_le_bytes());
            payload.extend_from_slice(&itab_last.to_le_bytes());
        }
        payload
    }

    #[test]
    fn parses_externsheet_entries_biff8() {
        let entries = [(1, 2, 3), (4, 5, 6)];
        let payload = externsheet_payload(&entries);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed =
            parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            parsed,
            vec![
                ExternSheetEntry {
                    supbook: 1,
                    itab_first: 2,
                    itab_last: 3
                },
                ExternSheetEntry {
                    supbook: 4,
                    itab_first: 5,
                    itab_last: 6
                }
            ]
        );
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

        let parsed =
            parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(
            parsed,
            vec![
                ExternSheetEntry {
                    supbook: 0x0010,
                    itab_first: 0x0020,
                    itab_last: 0x0030
                },
                ExternSheetEntry {
                    supbook: 0x0040,
                    itab_first: 0x0050,
                    itab_last: 0x0060
                }
            ]
        );
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

        let parsed =
            parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].itab_first, 2);
        assert_eq!(parsed[1].itab_last, 6);
    }

    #[test]
    fn scan_stops_on_malformed_record_and_returns_partial() {
        let entries = [(1, 2, 3), (4, 5, 6)];
        let payload = externsheet_payload(&entries);

        // Truncated record: declares 4 bytes but only provides 2. Must be at end of stream.
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x1234u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[0xAA, 0xBB]);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_EXTERNSHEET, &payload),
            truncated,
        ]
        .concat();

        let parsed =
            parse_biff_externsheet(&stream, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].supbook, 1);
        assert_eq!(parsed[1].itab_last, 6);
    }
}
