//! BIFF NOTE/OBJ/TXO parsing for legacy cell comments ("notes").
//!
//! Excel 97-2003 `.xls` files store cell notes as a small record graph:
//! - `NOTE`: the cell anchor (row/col) + the displayed author string
//! - `OBJ` (ftCmo): links the note to a drawing object id
//! - `TXO` (+ `CONTINUE` records): stores the comment text payload
//!
//! This parser is intentionally best-effort and lossy:
//! - Only plain text + author are decoded; rich text formatting and comment box
//!   geometry/visibility are ignored.
//! - Malformed/incomplete record sequences may yield partial output and warnings.
//! - Modern threaded comments are an OOXML feature and are not supported in `.xls`.
//! - Missing TXO payloads (text) are treated as a warning and the note may be
//!   skipped by the importer.

use std::collections::HashMap;

use formula_model::{CellRef, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

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
// - `cbRuns` lives at offset 12 (byte length of the rich-text formatting run data that follows
//   the character bytes in the TXO continuation area)
// - the record is followed by `CONTINUE` records containing the character bytes and formatting runs
const TXO_TEXT_LEN_OFFSET: usize = 6;
const TXO_RUNS_LEN_OFFSET: usize = 12;

const MAX_WARNINGS_PER_SHEET: usize = 20;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BiffNote {
    pub(crate) cell: CellRef,
    /// The drawing object id (`idObj`) that was used to resolve this note's TXO payload.
    ///
    /// NOTE records redundantly store two adjacent u16 fields (`grbit` + `idObj`) and some
    /// producers appear to swap the ordering. We keep both candidates during parsing and then
    /// choose whichever one has a matching TXO record. This field captures that chosen id so
    /// callers can derive stable identifiers for imported comments.
    pub(crate) obj_id: u16,
    pub(crate) author: String,
    pub(crate) text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub(crate) struct ParsedSheetNotes {
    pub(crate) notes: Vec<BiffNote>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ParsedNote {
    cell: CellRef,
    primary_obj_id: u16,
    secondary_obj_id: u16,
    author: String,
}

/// Best-effort parse of legacy note comments from a worksheet BIFF substream.
///
/// Returns parsed notes + non-fatal warnings (bounded per sheet) so callers can surface partial
/// `.xls` imports to users and aid debugging.
pub(crate) fn parse_biff_sheet_notes(
    workbook_stream: &[u8],
    start: usize,
    biff: BiffVersion,
    codepage: u16,
) -> Result<ParsedSheetNotes, String> {
    let allows_continuation = |record_id: u16| record_id == RECORD_TXO;
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    let mut notes: Vec<ParsedNote> = Vec::new();
    let mut texts_by_obj_id: HashMap<u16, String> = HashMap::new();
    let mut current_obj_id: Option<u16> = None;
    let mut warnings: Vec<String> = Vec::new();

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                // Best-effort: stop on a malformed record, but surface a warning so callers can
                // report partial import to the user.
                push_warning(&mut warnings, format!("malformed BIFF record: {err}"));
                break;
            }
        };

        if record.offset != start && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_NOTE => {
                if let Some(note) =
                    parse_note_record(record.data.as_ref(), record.offset, biff, codepage, &mut warnings)
                {
                    notes.push(note);
                }
            }
            RECORD_OBJ => {
                current_obj_id =
                    parse_obj_record_id(record.data.as_ref(), record.offset, &mut warnings);
            }
            RECORD_TXO => {
                // The current object id applies only to the next TXO record.
                let Some(obj_id) = current_obj_id.take() else {
                    push_warning(
                        &mut warnings,
                        format!(
                            "TXO record at offset {} missing preceding OBJ object id",
                            record.offset
                        ),
                    );
                    continue;
                };

                if let Some(text) = parse_txo_text(&record, biff, codepage, &mut warnings) {
                    texts_by_obj_id.insert(obj_id, text);
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let mut out: Vec<BiffNote> = Vec::with_capacity(notes.len());
    let mut out_by_obj_id: HashMap<u16, usize> = HashMap::new();
    for note in notes {
        let Some((obj_id, text)) = texts_by_obj_id
            .get(&note.primary_obj_id)
            .map(|text| (note.primary_obj_id, text))
            .or_else(|| texts_by_obj_id.get(&note.secondary_obj_id).map(|text| (note.secondary_obj_id, text)))
        else {
            // No TXO payload for this NOTE record: keep best-effort import going, but skip creating
            // a model comment with missing text.
            push_warning(
                &mut warnings,
                format!(
                    "NOTE record for cell {} references missing TXO payload (obj_id={}, fallback_obj_id={})",
                    note.cell.to_a1(),
                    note.primary_obj_id,
                    note.secondary_obj_id
                ),
            );
            continue;
        };

        let resolved = BiffNote {
            cell: note.cell,
            obj_id,
            author: note.author,
            text: text.clone(),
        };

        if let Some(&existing) = out_by_obj_id.get(&obj_id) {
            warnings.push(format!(
                "duplicate NOTE record for object id {obj_id} (cell {}); overwriting previous NOTE at cell {}",
                resolved.cell.to_a1(),
                out.get(existing)
                    .map(|note| note.cell.to_a1())
                    .unwrap_or_else(|| "<unknown>".to_string())
            ));
            if let Some(slot) = out.get_mut(existing) {
                *slot = resolved;
            }
        } else {
            out_by_obj_id.insert(obj_id, out.len());
            out.push(resolved);
        }
    }

    Ok(ParsedSheetNotes {
        notes: out,
        warnings,
    })
}

fn push_warning(warnings: &mut Vec<String>, warning: impl Into<String>) {
    if warnings.len() >= MAX_WARNINGS_PER_SHEET {
        return;
    }
    warnings.push(warning.into());
}

fn parse_note_record(
    data: &[u8],
    offset: usize,
    biff: BiffVersion,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<ParsedNote> {
    if data.len() < 8 {
        push_warning(
            warnings,
            format!(
                "NOTE record at offset {offset} is too short (len={})",
                data.len()
            ),
        );
        return None;
    }

    let row = u16::from_le_bytes([data[0], data[1]]) as u32;
    let col = u16::from_le_bytes([data[2], data[3]]) as u32;
    // Some parsers differ on whether `idObj` precedes `grbit`. Capture both fields and match them
    // up with OBJ/TXO payloads later (join by object id).
    let primary_obj_id = u16::from_le_bytes([data[6], data[7]]);
    let secondary_obj_id = u16::from_le_bytes([data[4], data[5]]);
    if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
        push_warning(
            warnings,
            format!(
                "NOTE record at offset {offset} references out-of-bounds cell ({row},{col}) (obj_id={primary_obj_id}, fallback_obj_id={secondary_obj_id})"
            ),
        );
        return None;
    }

    // `stAuthor` is specified as a `ShortXLUnicodeString` (BIFF8) or an ANSI short string (BIFF5),
    // but files in the wild sometimes store it as an `XLUnicodeString` (16-bit length prefix), and
    // may include embedded NULs. Keep this best-effort: return an empty author string if decoding
    // fails.
    let author_bytes = &data[8..];
    let mut author = match strings::parse_biff_short_string(author_bytes, biff, codepage) {
        Ok((s, consumed)) => {
            // If BIFF8 short-string parsing succeeds but doesn't consume the full payload, attempt
            // `XLUnicodeString` decoding before falling back to the short-string result.
            if biff == BiffVersion::Biff8 && consumed != author_bytes.len() {
                match strings::parse_biff8_unicode_string(author_bytes, codepage) {
                    Ok((alt, alt_consumed)) if alt_consumed == author_bytes.len() => alt,
                    _ => s,
                }
            } else {
                s
            }
        }
        Err(err) => {
            if biff == BiffVersion::Biff8 {
                match strings::parse_biff8_unicode_string(author_bytes, codepage) {
                    Ok((alt, _)) => alt,
                    Err(unicode_err) => {
                        push_warning(
                            warnings,
                            format!(
                                "failed to parse NOTE author string at offset {offset}: {err}; XLUnicodeString fallback also failed: {unicode_err}"
                            ),
                        );
                        String::new()
                    }
                }
            } else {
                push_warning(
                    warnings,
                    format!("failed to parse NOTE author string at offset {offset}: {err}"),
                );
                String::new()
            }
        }
    };
    strip_embedded_nuls(&mut author);

    Some(ParsedNote {
        cell: CellRef::new(row, col),
        primary_obj_id,
        secondary_obj_id,
        author,
    })
}

fn parse_obj_record_id(
    data: &[u8],
    record_offset: usize,
    warnings: &mut Vec<String>,
) -> Option<u16> {
    let mut idx = 0usize;
    let mut obj_id: Option<u16> = None;

    while idx + 4 <= data.len() {
        let ft = u16::from_le_bytes([data[idx], data[idx + 1]]);
        let cb = u16::from_le_bytes([data[idx + 2], data[idx + 3]]) as usize;
        idx += 4;

        let end = match idx.checked_add(cb) {
            Some(end) => end,
            None => {
                push_warning(
                    warnings,
                    format!("OBJ record at offset {record_offset} has subrecord length overflow"),
                );
                break;
            }
        };
        let Some(sub) = data.get(idx..end) else {
            push_warning(
                warnings,
                "OBJ record at offset {record_offset} has truncated subrecord 0x{ft:04X} (cb={cb})"
            );
            break;
        };

        if ft == OBJ_SUBRECORD_FT_CMO {
            // ftCmo: ot (2) + id (2) + ...
            if sub.len() >= 4 {
                obj_id = Some(u16::from_le_bytes([sub[2], sub[3]]));
            } else {
                push_warning(
                    warnings,
                    format!(
                        "OBJ record has truncated ftCmo subrecord (len={})",
                        sub.len()
                    ),
                );
            }
        }

        idx = end;
    }

    if obj_id.is_none() {
        push_warning(
            warnings,
            format!("OBJ record at offset {record_offset} missing ftCmo object id"),
        );
    }
    obj_id
}

fn parse_txo_text_with_warnings(
    record: &records::LogicalBiffRecord<'_>,
    biff: BiffVersion,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    match biff {
        BiffVersion::Biff5 => parse_txo_text_biff5(record, codepage, warnings),
        BiffVersion::Biff8 => parse_txo_text_biff8(record, codepage, warnings),
    }
}

fn parse_txo_text_biff5(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    // BIFF5 typically stores the TXO text bytes directly in subsequent CONTINUE records (no
    // per-fragment option flags byte). Treat the continued bytes as ANSI encoded using the
    // workbook codepage.
    //
    // Some producers appear to mimic BIFF8's continued-string layout and prefix each CONTINUE
    // fragment with a one-byte "high-byte" flag (0/1). In that case, the TXO `cchText` count does
    // *not* include those flag bytes, so treat them as optional and skip them best-effort.
    let first = record.first_fragment();
    if first.len() < TXO_TEXT_LEN_OFFSET + 2 {
        // Best-effort: if the TXO header is malformed, decode the first continuation as text.
        return fallback_decode_first_continue_biff5(record, codepage, warnings);
    }
    let cch_text = u16::from_le_bytes([
        first[TXO_TEXT_LEN_OFFSET],
        first[TXO_TEXT_LEN_OFFSET + 1],
    ]) as usize;
    if cch_text == 0 {
        return Some(String::new());
    }

    let fragments: Vec<&[u8]> = record.fragments().collect();
    let continues = fragments.get(1..).unwrap_or_default();
    if continues.is_empty() {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} missing CONTINUE fragments (expected {cch_text} chars)",
                record.offset
            ),
        );
        return Some(String::new());
    }

    let mut skip_leading_flag_bytes = false;
    if let Some(first_continue) = continues.first().copied() {
        if matches!(first_continue.first().copied(), Some(0) | Some(1)) {
            let max_chars_if_flags_present: usize = continues
                .iter()
                .map(|frag| {
                    if frag.is_empty() {
                        0
                    } else if matches!(frag.first().copied(), Some(0) | Some(1)) {
                        frag.len().saturating_sub(1)
                    } else {
                        frag.len()
                    }
                })
                .sum();

            if cch_text <= max_chars_if_flags_present {
                skip_leading_flag_bytes = true;
            }
        }
    }

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

        if offset == 0 && skip_leading_flag_bytes && matches!(frag.first().copied(), Some(0) | Some(1))
        {
            offset = 1;
        }

        if offset >= frag.len() {
            frag_idx += 1;
            offset = 0;
            continue;
        }

        let available = frag.len() - offset;
        let take = remaining.min(available);
        out.push_str(&strings::decode_ansi(codepage, &frag[offset..offset + take]));
        remaining -= take;
        offset += take;
        if offset >= frag.len() {
            frag_idx += 1;
            offset = 0;
        }
    }

    if remaining > 0 {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} truncated text (wanted {cch_text} chars, got {})",
                record.offset,
                cch_text.saturating_sub(remaining)
            ),
        );
    }
    trim_trailing_nuls(&mut out);
    strip_embedded_nuls(&mut out);
    Some(out)
}

fn parse_txo_text_biff8(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    let first = record.first_fragment();
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let continues = fragments.get(1..).unwrap_or_default();
    if continues.is_empty() {
        match parse_txo_cch_text(first, 0) {
            Some(0) => {}
            Some(cch_text) => {
                push_warning(
                    warnings,
                    format!(
                        "TXO record at offset {} missing CONTINUE fragments (expected {cch_text} chars)",
                        record.offset
                    ),
                );
            }
            None => {
                push_warning(
                    warnings,
                    format!(
                        "TXO record at offset {} missing CONTINUE fragments (unable to read cchText from header)",
                        record.offset
                    ),
                );
            }
        }
        return Some(String::new());
    }

    // `cbRuns` indicates how many bytes at the end of the TXO continuation area are reserved for
    // rich-text formatting runs. We ignore those bytes so we don't misinterpret formatting run data
    // as characters if `cchText` is larger than the available text bytes (truncated/corrupt files).
    let cb_runs = first
        .get(TXO_RUNS_LEN_OFFSET..TXO_RUNS_LEN_OFFSET + 2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]) as usize)
        .unwrap_or(0);
    let total_continue_bytes: usize = continues.iter().map(|frag| frag.len()).sum();
    let text_continue_bytes = total_continue_bytes.saturating_sub(cb_runs);
    if cb_runs > total_continue_bytes {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} has cbRuns ({cb_runs}) larger than continuation payload ({total_continue_bytes})",
                record.offset
            ),
        );
    }

    let max_chars = estimate_max_chars_with_byte_limit(continues, text_continue_bytes);
    let Some(cch_text) = parse_txo_cch_text(first, max_chars) else {
        return fallback_decode_first_continue(record, codepage, warnings);
    };
    if cch_text == 0 {
        return Some(String::new());
    }

    let mut out = String::new();
    let mut remaining = cch_text;
    let mut remaining_bytes = text_continue_bytes;
    for frag in continues {
        if remaining == 0 || remaining_bytes == 0 {
            break;
        }

        let take_len = frag.len().min(remaining_bytes);
        remaining_bytes = remaining_bytes.saturating_sub(take_len);
        let frag = &frag[..take_len];
        let Some((&flags, bytes)) = frag.split_first() else {
            continue;
        };
        let is_unicode = (flags & 0x01) != 0;
        let bytes_per_char = if is_unicode { 2 } else { 1 };
        let available_chars = bytes.len() / bytes_per_char;
        if available_chars == 0 {
            continue;
        }

        let take_chars = remaining.min(available_chars);
        let take_bytes = take_chars * bytes_per_char;
        let slice = &bytes[..take_bytes];
        if is_unicode {
            out.push_str(&decode_utf16le(slice));
        } else {
            out.push_str(&strings::decode_ansi(codepage, slice));
        }
        remaining -= take_chars;
    }

    if remaining > 0 {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} truncated text (wanted {cch_text} chars, got {})",
                record.offset,
                cch_text.saturating_sub(remaining)
            ),
        );
    }
    trim_trailing_nuls(&mut out);
    strip_embedded_nuls(&mut out);
    Some(out)
}

fn parse_txo_text(
    record: &records::LogicalBiffRecord<'_>,
    biff: BiffVersion,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    parse_txo_text_with_warnings(record, biff, codepage, warnings)
}

fn parse_txo_cch_text(header: &[u8], max_chars: usize) -> Option<usize> {
    // Heuristic: cchText is typically at offset 6 in the TXO header, but some sources disagree.
    // Try a few common offsets and choose the first plausible value.
    if header.len() < 8 {
        return None;
    }

    // Spec-defined BIFF8 offset for cchText.
    let cch_at_6 =
        u16::from_le_bytes([header[TXO_TEXT_LEN_OFFSET], header[TXO_TEXT_LEN_OFFSET + 1]]) as usize;

    // If we have no continuation bytes to sanity-check against, trust the header.
    if max_chars == 0 {
        return Some(cch_at_6);
    }

    // In truncated/corrupt files we may not have enough continuation bytes to satisfy `cchText`.
    // In that case, prefer the spec-defined header field and let the decoder emit a truncation
    // warning rather than guessing a different offset that happens to fit the observed payload.
    if cch_at_6 != 0 {
        return Some(cch_at_6);
    }

    // If the spec-defined field is zero but we have continuation bytes, the header may be
    // malformed or use a non-standard layout. Fall back to scanning a few other offsets for a
    // plausible length.
    for off in [4usize, 8usize, 10usize, 12usize] {
        if header.len() < off + 2 {
            continue;
        }
        let cch = u16::from_le_bytes([header[off], header[off + 1]]) as usize;
        if cch != 0 && cch <= max_chars {
            return Some(cch);
        }
    }

    // Last resort: decode as much as we can from the available continuation bytes.
    Some(max_chars)
}

fn estimate_max_chars_with_byte_limit(continues: &[&[u8]], byte_limit: usize) -> usize {
    // Best-effort estimate used only for header heuristics.
    let mut total = 0usize;
    let mut remaining = byte_limit;
    for frag in continues {
        if remaining == 0 {
            break;
        }
        let take_len = frag.len().min(remaining);
        remaining = remaining.saturating_sub(take_len);
        let frag = &frag[..take_len];
        let Some((&flags, bytes)) = frag.split_first() else {
            continue;
        };
        let bytes_per_char = if (flags & 0x01) != 0 { 2 } else { 1 };
        total = total.saturating_add(bytes.len() / bytes_per_char);
    }
    total
}

fn fallback_decode_first_continue(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    let mut fragments = record.fragments();
    let _ = fragments.next(); // skip header
    let Some(first) = fragments.next() else {
        push_warning(
            warnings,
            format!("TXO record at offset {} missing CONTINUE fragments", record.offset),
        );
        return Some(String::new());
    };

    push_warning(
        warnings,
        format!(
            "TXO record at offset {} has malformed header; falling back to decoding first CONTINUE fragment",
            record.offset
        ),
    );
    let Some((&flags, bytes)) = first.split_first() else {
        return Some(String::new());
    };
    let mut out = if (flags & 0x01) != 0 {
        decode_utf16le(bytes)
    } else {
        strings::decode_ansi(codepage, bytes)
    };
    trim_trailing_nuls(&mut out);
    strip_embedded_nuls(&mut out);
    Some(out)
}

fn fallback_decode_first_continue_biff5(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    let mut fragments = record.fragments();
    let _ = fragments.next(); // skip header
    let Some(first) = fragments.next() else {
        push_warning(
            warnings,
            format!("TXO record at offset {} missing CONTINUE fragments", record.offset),
        );
        return Some(String::new());
    };

    push_warning(
        warnings,
        format!(
            "TXO record at offset {} has malformed header; falling back to decoding first CONTINUE fragment",
            record.offset
        ),
    );

    if first.is_empty() {
        return Some(String::new());
    }

    // BIFF5 usually stores the raw ANSI bytes directly, but some producers mimic the BIFF8
    // continued-string layout and prefix the fragment with a 0/1 "high-byte" flag.
    let mut out = if matches!(first.first().copied(), Some(0) | Some(1)) {
        let flags = first[0];
        let bytes = &first[1..];
        if (flags & 0x01) != 0 {
            decode_utf16le(bytes)
        } else {
            strings::decode_ansi(codepage, bytes)
        }
    } else {
        strings::decode_ansi(codepage, first)
    };

    trim_trailing_nuls(&mut out);
    strip_embedded_nuls(&mut out);
    Some(out)
}

fn decode_utf16le(bytes: &[u8]) -> String {
    let mut u16s = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    String::from_utf16_lossy(&u16s)
}

fn strip_embedded_nuls(s: &mut String) {
    if s.contains('\0') {
        s.retain(|c| c != '\0');
    }
}

fn trim_trailing_nuls(s: &mut String) {
    while s.chars().last() == Some('\0') {
        s.pop();
    }
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

    fn bof_biff5() -> Vec<u8> {
        record(records::RECORD_BOF_BIFF5, &[0u8; 16])
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

    fn note_with_xl_unicode_author(row: u16, col: u16, obj_id: u16, author: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        payload.extend_from_slice(&obj_id.to_le_bytes());
        payload.extend_from_slice(&obj_id.to_le_bytes());

        // BIFF8 XLUnicodeString author (u16 length).
        payload.extend_from_slice(&(author.len() as u16).to_le_bytes());
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(author.as_bytes());

        record(RECORD_NOTE, &payload)
    }

    fn note_biff5(row: u16, col: u16, obj_id: u16, author: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&obj_id.to_le_bytes()); // idObj

        // BIFF5 short ANSI string (length + bytes).
        payload.push(author.len() as u8);
        payload.extend_from_slice(author.as_bytes());

        record(RECORD_NOTE, &payload)
    }

    fn note_biff5_author_bytes(row: u16, col: u16, obj_id: u16, author_bytes: &[u8]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&row.to_le_bytes());
        payload.extend_from_slice(&col.to_le_bytes());
        // NOTE record stores `grbit` and `idObj` as two adjacent u16 fields; the ordering varies
        // across parsers, so we write the same value into both fields to keep the fixture robust.
        payload.extend_from_slice(&obj_id.to_le_bytes());
        payload.extend_from_slice(&obj_id.to_le_bytes());

        payload.push(author_bytes.len() as u8);
        payload.extend_from_slice(author_bytes);
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
        // TXO header with `cchText` at offset 6.
        let mut payload = vec![0u8; 18];
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&(text.len() as u16).to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn txo_with_cch_text(cch_text: u16) -> Vec<u8> {
        let mut payload = vec![0u8; 18];
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2].copy_from_slice(&cch_text.to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn txo_with_cch_text_and_cb_runs(cch_text: u16, cb_runs: u16) -> Vec<u8> {
        let mut payload = vec![0u8; 18];
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2].copy_from_slice(&cch_text.to_le_bytes());
        payload[TXO_RUNS_LEN_OFFSET..TXO_RUNS_LEN_OFFSET + 2].copy_from_slice(&cb_runs.to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn txo_with_cch_text_at_offset_4(cch_text: u16) -> Vec<u8> {
        // Some sources disagree on the TXO header layout. This helper intentionally writes
        // `cchText` at offset 4 instead of the spec-defined offset 6.
        let mut payload = vec![0u8; 18];
        payload[4..6].copy_from_slice(&cch_text.to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn continue_text_ascii(text: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.push(0); // fHighByte=0 (compressed 8-bit)
        payload.extend_from_slice(text.as_bytes());
        record(records::RECORD_CONTINUE, &payload)
    }

    fn continue_text_compressed_bytes(bytes: &[u8]) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.push(0); // fHighByte=0 (compressed 8-bit)
        payload.extend_from_slice(bytes);
        record(records::RECORD_CONTINUE, &payload)
    }

    fn continue_text_biff5(bytes: &[u8]) -> Vec<u8> {
        record(records::RECORD_CONTINUE, bytes)
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
    fn note_record_strips_embedded_nuls_from_author() {
        let author = "Al\0ice";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // row
        payload.extend_from_slice(&0u16.to_le_bytes()); // col
        payload.extend_from_slice(&1u16.to_le_bytes()); // grbit/idObj
        payload.extend_from_slice(&1u16.to_le_bytes()); // idObj/grbit

        // BIFF8 ShortXLUnicodeString author (compressed) containing an embedded NUL.
        payload.push(author.len() as u8);
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(author.as_bytes());

        let mut warnings = Vec::new();
        let parsed = parse_note_record(&payload, 0, BiffVersion::Biff8, 1252, &mut warnings)
            .expect("parse note");
        assert_eq!(parsed.author, "Alice");
    }

    #[test]
    fn note_record_parses_author_encoded_as_xlunicode_string() {
        let author = "Alice";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // row
        payload.extend_from_slice(&0u16.to_le_bytes()); // col
        payload.extend_from_slice(&1u16.to_le_bytes()); // grbit/idObj
        payload.extend_from_slice(&1u16.to_le_bytes()); // idObj/grbit

        // BIFF8 XLUnicodeString author (16-bit length).
        payload.extend_from_slice(&(author.len() as u16).to_le_bytes());
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(author.as_bytes());

        let mut warnings = Vec::new();
        let parsed = parse_note_record(&payload, 0, BiffVersion::Biff8, 1252, &mut warnings)
            .expect("parse note");
        assert_eq!(parsed.author, author);
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

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        let note = &notes[0];
        assert_eq!(note.cell, CellRef::new(0, 0));
        assert_eq!(note.obj_id, 1);
        assert_eq!(note.author, "Alice");
        assert_eq!(note.text, "Hello");
    }

    #[test]
    fn parses_biff5_txo_text_from_continue_without_flags_using_codepage() {
        // Include a non-ASCII byte in the BIFF5 text (0xC0 => Cyrillic 'А' in Windows-1251) to
        // ensure codepage decoding is applied.
        let text_bytes = [b'H', b'i', b' ', 0xC0];

        let mut txo_payload = vec![0u8; 18];
        txo_payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&(text_bytes.len() as u16).to_le_bytes());

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_payload),
            continue_text_biff5(&text_bytes),
            // Formatting CONTINUE (ignored).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1251).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hi А");
    }

    #[test]
    fn parses_biff5_note_author_using_codepage() {
        // Windows-1251 0xC0 => Cyrillic 'А' (U+0410). This ensures the BIFF5 short ANSI author
        // string goes through `strings::decode_ansi` using the workbook codepage.
        let author_bytes = [0xC0];

        let stream = [
            bof_biff5(),
            note_biff5_author_bytes(0, 0, 1, &author_bytes),
            obj_with_id(1),
            txo_with_cch_text(2),
            continue_text_biff5(b"Hi"),
            // Formatting CONTINUE (ignored).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1251).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "А");
        assert_eq!(notes[0].text, "Hi");
    }

    #[test]
    fn strips_embedded_nuls_from_author() {
        let stream = [
            bof(),
            note(0, 0, 1, "Al\0ice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
    }

    #[test]
    fn trims_trailing_nuls_from_text() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello\0"),
            continue_text_ascii("Hello\0"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn strips_embedded_nuls_from_text() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("H\0e\0l\0l\0o"),
            continue_text_ascii("H\0e\0l\0l\0o"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
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

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 2);

        let mut by_cell: HashMap<CellRef, &BiffNote> = HashMap::new();
        for note in &notes {
            by_cell.insert(note.cell, note);
        }

        let n1 = by_cell.get(&CellRef::new(0, 0)).expect("note 1");
        assert_eq!(n1.cell, CellRef::new(0, 0));
        assert_eq!(n1.obj_id, 1);
        assert_eq!(n1.author, "Alice");
        assert_eq!(n1.text, "First");

        let n2 = by_cell.get(&CellRef::new(1, 1)).expect("note 2");
        assert_eq!(n2.cell, CellRef::new(1, 1));
        assert_eq!(n2.obj_id, 2);
        assert_eq!(n2.author, "Bob");
        assert_eq!(n2.text, "Second");
    }

    #[test]
    fn dedupes_duplicate_note_records_by_object_id() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            // Duplicate NOTE for the same object id; later record should win.
            note(1, 1, 1, "Bob"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].cell, CellRef::new(1, 1));
        assert_eq!(notes[0].author, "Bob");
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("duplicate NOTE record")),
            "expected duplicate NOTE warning; warnings={warnings:?}"
        );
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

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
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

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            !warnings.is_empty(),
            "expected warnings for truncated record"
        );
    }

    #[test]
    fn skips_out_of_bounds_note_columns() {
        let stream = [
            bof(),
            // NOTE references an out-of-bounds column (col=EXCEL_MAX_COLS).
            note(0, EXCEL_MAX_COLS as u16, 1, "Alice"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(notes.is_empty(), "expected note to be skipped");
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("out-of-bounds") && w.contains("obj_id=1")),
            "expected out-of-bounds warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn falls_back_to_decoding_first_continue_when_txo_header_is_missing() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            // Empty TXO header.
            record(RECORD_TXO, &[]),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("falling back")),
            "expected fallback warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn falls_back_to_decoding_first_continue_when_txo_header_is_missing_biff5() {
        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            // Empty TXO header.
            record(RECORD_TXO, &[]),
            continue_text_biff5(b"Hi"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hi");
        assert!(
            warnings.iter().any(|w| w.contains("falling back")),
            "expected fallback warning; warnings={warnings:?}"
        );
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

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hi");
    }

    #[test]
    fn parses_compressed_text_from_continue_using_codepage() {
        // In Windows-1251, 0xC0 is Cyrillic 'А' (U+0410); in Windows-1252 it's 'À'.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(1),
            continue_text_compressed_bytes(&[0xC0]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1251).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "\u{0410}");
    }

    #[test]
    fn parses_text_split_across_multiple_continue_records() {
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(5),
            continue_text_compressed_bytes(b"AB"),
            continue_text_compressed_bytes(b"CDE"),
            // Formatting runs CONTINUE payload (dummy bytes).
            record(records::RECORD_CONTINUE, &[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "ABCDE");
    }

    #[test]
    fn does_not_decode_formatting_run_bytes_as_text_when_cbruns_is_set() {
        // cchText is intentionally larger than the available text bytes. Without respecting
        // `cbRuns`, parsers may decode the formatting run bytes as if they were characters.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            // cchText=6 but only 5 chars of actual text bytes.
            txo_with_cch_text_and_cb_runs(6, 4),
            continue_text_ascii("Hello"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0xFFu8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_txo_text_when_cchtext_is_stored_at_alternate_offset() {
        // The spec-defined cchText field (offset 6) is zero, but offset 4 contains the correct
        // value. Best-effort decoding should still recover the full text.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_at_offset_4(5),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn parses_biff5_txo_text_split_across_multiple_continue_records_without_flags_using_codepage() {
        // Split the BIFF5 text across two CONTINUE records to validate concatenation across record
        // boundaries. Include a non-ASCII byte so Windows-1251 decoding is exercised (0xC0 => 'А').
        let part1 = [b'H', b'i', b' '];
        let part2 = [0xC0];
        let cch_text = (part1.len() + part2.len()) as u16;

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(cch_text),
            continue_text_biff5(&part1),
            continue_text_biff5(&part2),
            // Formatting CONTINUE payload (dummy bytes).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1251).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hi А");
    }

    #[test]
    fn parses_biff5_txo_text_split_across_multiple_continue_records_with_flags_using_codepage() {
        // Some BIFF5 writers appear to prefix each CONTINUE fragment with a BIFF8-style
        // "high-byte" flag (0/1). Ensure we treat that as an optional flag byte rather than part of
        // the text payload.
        let part1 = [b'H', b'i', b' '];
        let part2 = [0xC0];
        let cch_text = (part1.len() + part2.len()) as u16;

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(cch_text),
            // Use the BIFF8-style helper so each fragment begins with a flag byte.
            continue_text_compressed_bytes(&part1),
            continue_text_compressed_bytes(&part2),
            // Formatting CONTINUE payload (dummy bytes).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1251).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hi А");
    }

    #[test]
    fn parses_note_author_as_xl_unicode_string_when_short_string_leaves_trailing_bytes() {
        // Some BIFF8 producers store NOTE authors as XLUnicodeString (u16 length) instead of
        // ShortXLUnicodeString (u8 length). Our parser should detect this and fall back.
        let stream = [
            bof(),
            note_with_xl_unicode_author(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
    }
}
