#![allow(dead_code)]

use std::collections::HashMap;

use formula_model::CellRef;

use super::{records, strings, BiffVersion};

// Worksheet record ids used to recover legacy Excel "notes" (cell comments).
// See [MS-XLS]:
// - NOTE: 2.4.168
// - OBJ: 2.4.163
// - TXO: 2.4.334
const RECORD_NOTE: u16 = 0x001C;
const RECORD_OBJ: u16 = 0x005D;
const RECORD_TXO: u16 = 0x01B6;

// OBJ subrecord types. We only need `ftCmo`, which includes the drawing object's id.
// See [MS-XLS] 2.5.49 (ftCmo).
const OBJ_SUBRECORD_FT_CMO: u16 = 0x0015;

// TXO record payload layout [MS-XLS 2.4.334]:
// - `cchText` lives at offset 6
// - the record is followed by `CONTINUE` records containing the character bytes and formatting runs
const TXO_TEXT_LEN_OFFSET: usize = 6;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct SheetNote {
    pub(crate) cell: CellRef,
    pub(crate) obj_id: u16,
    pub(crate) author: String,
    pub(crate) text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ParsedNote {
    cell: CellRef,
    primary_obj_id: u16,
    secondary_obj_id: u16,
    author: String,
}

pub(crate) fn parse_biff_sheet_notes(
    workbook_stream: &[u8],
    start: usize,
    biff: BiffVersion,
    codepage: u16,
) -> Result<Vec<SheetNote>, String> {
    let allows_continuation = |record_id: u16| record_id == RECORD_TXO;
    let iter = records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    let mut notes: Vec<ParsedNote> = Vec::new();
    let mut texts_by_obj_id: HashMap<u16, String> = HashMap::new();
    let mut current_obj_id: Option<u16> = None;

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(_) => break, // best-effort: stop on malformed record
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_NOTE => {
                if let Some(note) = parse_note_record(record.data.as_ref(), biff, codepage) {
                    notes.push(note);
                }
            }
            RECORD_OBJ => {
                current_obj_id = parse_obj_record_id(record.data.as_ref());
            }
            RECORD_TXO => {
                if let Some(obj_id) = current_obj_id {
                    if let Some(text) = parse_txo_text(&record, biff, codepage) {
                        texts_by_obj_id.insert(obj_id, text);
                    }
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let mut out = Vec::with_capacity(notes.len());
    for note in notes {
        let obj_id = if texts_by_obj_id.contains_key(&note.primary_obj_id) {
            note.primary_obj_id
        } else if texts_by_obj_id.contains_key(&note.secondary_obj_id) {
            note.secondary_obj_id
        } else {
            note.primary_obj_id
        };

        let text = texts_by_obj_id.get(&obj_id).cloned().unwrap_or_default();
        out.push(SheetNote {
            cell: note.cell,
            obj_id,
            author: note.author,
            text,
        });
    }

    Ok(out)
}

fn parse_note_record(data: &[u8], biff: BiffVersion, codepage: u16) -> Option<ParsedNote> {
    if data.len() < 8 {
        return None;
    }

    let row = u16::from_le_bytes([data[0], data[1]]) as u32;
    let col = u16::from_le_bytes([data[2], data[3]]) as u32;
    // Some parsers differ on whether `idObj` precedes `grbit`. Capture both fields and match them
    // up with OBJ/TXO payloads later (join by object id).
    let primary_obj_id = u16::from_le_bytes([data[6], data[7]]);
    let secondary_obj_id = u16::from_le_bytes([data[4], data[5]]);

    let author = strings::parse_biff_short_string(&data[8..], biff, codepage)
        .map(|(s, _)| s)
        .unwrap_or_default();

    Some(ParsedNote {
        cell: CellRef::new(row, col),
        primary_obj_id,
        secondary_obj_id,
        author,
    })
}

fn parse_obj_record_id(data: &[u8]) -> Option<u16> {
    let mut offset = 0usize;

    while offset + 4 <= data.len() {
        let ft = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let cb = u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as usize;
        offset += 4;

        let sub = data.get(offset..offset + cb)?;
        if ft == OBJ_SUBRECORD_FT_CMO && sub.len() >= 4 {
            // ftCmo: ot (2) + id (2) + ...
            return Some(u16::from_le_bytes([sub[2], sub[3]]));
        }

        offset = offset.checked_add(cb)?;
    }

    None
}

fn parse_txo_text(
    record: &records::LogicalBiffRecord<'_>,
    biff: BiffVersion,
    codepage: u16,
) -> Option<String> {
    match biff {
        BiffVersion::Biff5 => parse_txo_text_biff5(record, codepage),
        BiffVersion::Biff8 => parse_txo_text_biff8(record, codepage),
    }
}

fn parse_txo_text_biff5(record: &records::LogicalBiffRecord<'_>, codepage: u16) -> Option<String> {
    // BIFF5 notes are rare for us; keep this best-effort and only support the same continuation
    // layout as BIFF8, but decode 8-bit bytes using the workbook codepage.
    let first = record.first_fragment();
    if first.len() < TXO_TEXT_LEN_OFFSET + 2 {
        return Some(String::new());
    }
    let cch_text = u16::from_le_bytes([
        first[TXO_TEXT_LEN_OFFSET],
        first[TXO_TEXT_LEN_OFFSET + 1],
    ]) as usize;
    if cch_text == 0 {
        return Some(String::new());
    }

    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut frag_idx = 1usize;
    let mut offset = 0usize;
    let mut remaining = cch_text;
    let mut out = String::new();

    while remaining > 0 {
        let frag = fragments.get(frag_idx).copied().unwrap_or_default();
        if frag.is_empty() {
            frag_idx += 1;
            offset = 0;
            if frag_idx >= fragments.len() {
                break;
            }
            continue;
        }

        // Each fragment begins with a one-byte "high-byte" flag in BIFF8; for BIFF5 treat it as
        // reserved and always decode as ANSI.
        if offset == 0 {
            offset = 1;
        }

        let available = frag.len().saturating_sub(offset);
        if available == 0 {
            frag_idx += 1;
            offset = 0;
            continue;
        }
        let take = remaining.min(available);
        out.push_str(&strings::decode_ansi(codepage, &frag[offset..offset + take]));
        remaining -= take;
        offset += take;
        if offset >= frag.len() {
            frag_idx += 1;
            offset = 0;
        }
    }

    Some(out)
}

fn parse_txo_text_biff8(record: &records::LogicalBiffRecord<'_>, codepage: u16) -> Option<String> {
    let first = record.first_fragment();
    if first.len() < TXO_TEXT_LEN_OFFSET + 2 {
        return Some(String::new());
    }

    let cch_text = u16::from_le_bytes([
        first[TXO_TEXT_LEN_OFFSET],
        first[TXO_TEXT_LEN_OFFSET + 1],
    ]) as usize;
    if cch_text == 0 {
        return Some(String::new());
    }

    let fragments: Vec<&[u8]> = record.fragments().collect();
    // Fragment 0 is the TXO header; text lives in subsequent CONTINUE records.
    let mut frag_idx = 1usize;
    let mut offset = 0usize;
    let mut remaining = cch_text;
    let mut out = String::new();
    let mut is_unicode = false;

    while remaining > 0 {
        let frag = fragments.get(frag_idx).copied().unwrap_or_default();
        if frag.is_empty() {
            frag_idx += 1;
            offset = 0;
            if frag_idx >= fragments.len() {
                break;
            }
            continue;
        }

        if offset == 0 {
            // Each CONTINUE fragment begins with a one-byte "high-byte" flag (bit0) indicating
            // whether the following character data is UTF-16LE (1) or 8-bit compressed (0).
            is_unicode = (frag[0] & 0x01) != 0;
            offset = 1;
        }

        let bytes_per_char = if is_unicode { 2 } else { 1 };
        let available_bytes = frag.len().saturating_sub(offset);
        if available_bytes < bytes_per_char {
            // Truncated or split mid-character: move to next fragment (best-effort).
            frag_idx += 1;
            offset = 0;
            continue;
        }

        let available_chars = available_bytes / bytes_per_char;
        if available_chars == 0 {
            frag_idx += 1;
            offset = 0;
            continue;
        }

        let take_chars = remaining.min(available_chars);
        let take_bytes = take_chars * bytes_per_char;
        let bytes = &frag[offset..offset + take_bytes];

        if is_unicode {
            let mut u16s = Vec::with_capacity(take_chars);
            for chunk in bytes.chunks_exact(2) {
                u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
            }
            out.push_str(&String::from_utf16_lossy(&u16s));
        } else {
            out.push_str(&strings::decode_ansi(codepage, bytes));
        }

        remaining -= take_chars;
        offset += take_bytes;
        if offset >= frag.len() {
            frag_idx += 1;
            offset = 0;
        }
    }

    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    fn bof() -> Vec<u8> {
        record(records::RECORD_BOF_BIFF8, &[0u8; 16])
    }

    fn eof() -> Vec<u8> {
        record(records::RECORD_EOF, &[])
    }

    fn note(row: u16, col: u16, obj_id: u16, author: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        // NOTE record stores `grbit` and `idObj` as two adjacent u16 fields; the ordering varies
        // across parsers, so we write the same value into both to keep the fixture robust.
        payload.extend_from_slice(&obj_id.to_le_bytes());
        payload.extend_from_slice(&obj_id.to_le_bytes());

        // BIFF8 ShortXLUnicodeString author (compressed).
        payload.push(author.len() as u8);
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(author.as_bytes());

        record(RECORD_NOTE, &payload)
    }

    fn obj_with_id(obj_id: u16) -> Vec<u8> {
        // ftCmo subrecord:
        // - ft=0x0015
        // - cb=18
        // - ot (2) + id (2) + rest (14)
        let mut ftcmo = Vec::new();
        ftcmo.extend_from_slice(&OBJ_SUBRECORD_FT_CMO.to_le_bytes());
        ftcmo.extend_from_slice(&18u16.to_le_bytes());
        ftcmo.extend_from_slice(&0u16.to_le_bytes()); // ot (unused)
        ftcmo.extend_from_slice(&obj_id.to_le_bytes());
        ftcmo.extend_from_slice(&[0u8; 14]); // rest of ftCmo

        // ftEnd subrecord (optional).
        ftcmo.extend_from_slice(&0u16.to_le_bytes());
        ftcmo.extend_from_slice(&0u16.to_le_bytes());

        record(RECORD_OBJ, &ftcmo)
    }

    fn txo_with_text(text: &str) -> Vec<u8> {
        // TXO header with cchText at offset 4.
        let mut payload = vec![0u8; 18];
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&(text.len() as u16).to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn continue_text_ascii(text: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.push(0); // fHighByte=0 (compressed 8-bit)
        payload.extend_from_slice(text.as_bytes());
        record(records::RECORD_CONTINUE, &payload)
    }

    fn continue_text_unicode(text: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.push(0x01); // fHighByte=1 (UTF-16LE)
        for u in text.encode_utf16() {
            payload.extend_from_slice(&u.to_le_bytes());
        }
        record(records::RECORD_CONTINUE, &payload)
    }

    #[test]
    fn parses_single_note_obj_txo_text() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let notes = parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        let note = &notes[0];
        assert_eq!(note.cell, CellRef::new(0, 0));
        assert_eq!(note.obj_id, 1);
        assert_eq!(note.author, "Alice");
        assert_eq!(note.text, "Hello");
    }

    #[test]
    fn joins_note_and_text_by_obj_id() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            note(1, 1, 2, "Bob"),
            // OBJ/TXO for obj_id=2 comes first.
            obj_with_id(2),
            txo_with_text("Second"),
            continue_text_ascii("Second"),
            obj_with_id(1),
            txo_with_text("First"),
            continue_text_ascii("First"),
            eof(),
        ]
        .concat();

        let notes = parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 2);

        let mut by_id: HashMap<u16, &SheetNote> = HashMap::new();
        for note in &notes {
            by_id.insert(note.obj_id, note);
        }

        let n1 = by_id.get(&1).expect("note 1");
        assert_eq!(n1.cell, CellRef::new(0, 0));
        assert_eq!(n1.author, "Alice");
        assert_eq!(n1.text, "First");

        let n2 = by_id.get(&2).expect("note 2");
        assert_eq!(n2.cell, CellRef::new(1, 1));
        assert_eq!(n2.author, "Bob");
        assert_eq!(n2.text, "Second");
    }

    #[test]
    fn stops_at_next_bof() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            // Missing EOF for the first substream: second BOF starts a new substream.
            bof(),
            note(0, 1, 2, "Mallory"),
            obj_with_id(2),
            txo_with_text("ShouldNotParse"),
            continue_text_ascii("ShouldNotParse"),
            eof(),
        ]
        .concat();

        let notes = parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].obj_id, 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn best_effort_on_truncated_records() {
        let mut truncated = Vec::new();
        truncated.extend_from_slice(&0x1234u16.to_le_bytes());
        truncated.extend_from_slice(&4u16.to_le_bytes());
        truncated.extend_from_slice(&[0xAA, 0xBB]); // missing 2 bytes

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            truncated,
        ]
        .concat();

        let notes = parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn parses_unicode_text_from_continue() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hi"),
            continue_text_unicode("Hi"),
            eof(),
        ]
        .concat();

        let notes = parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hi");
    }
}
