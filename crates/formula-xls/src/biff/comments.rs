//! BIFF NOTE/OBJ/TXO parsing for legacy cell comments ("notes").
//!
//! Legacy `.xls` (BIFF5/BIFF8) files store cell notes as a small record graph:
//! - `NOTE`: the cell anchor (row/col) + the displayed author string
//! - `OBJ` (ftCmo): links the note to a drawing object id
//! - `TXO` (+ `CONTINUE` records): stores the comment text payload
//!
//! This parser is intentionally best-effort and lossy:
//! - Only plain text + author (when available) are decoded; rich text formatting
//!   and comment box geometry/visibility are ignored.
//! - Malformed/incomplete record sequences may yield partial output and warnings.
//! - Modern threaded comments are an OOXML feature and are not supported in `.xls`.
//! - Missing TXO payloads (text) are treated as a warning and the note may be
//!   skipped by the importer.
//!
//! BIFF5 vs BIFF8 notes:
//! - BIFF5 commonly stores `NOTE.stAuthor` as an ANSI short string (`CODEPAGE`), but some writers
//!   use BIFF8-style string encodings; we attempt to decode both best-effort.
//! - BIFF5 commonly stores TXO text continuation bytes as raw 8-bit codepage bytes (no per-fragment
//!   string option flags), but some writers prefix each fragment with a BIFF8-style 0/1 flag byte.
//!
//! Note: anchoring to merged regions (top-left cell) is handled by the importer
//! when inserting notes into the [`formula_model`] worksheet model.

use std::collections::{HashMap, HashSet};

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
const TXO_TEXT_LEN_OFFSETS: [usize; 4] = [TXO_TEXT_LEN_OFFSET, 4, 8, 10];
const TXO_MAX_TEXT_CHARS: usize = 32 * 1024;

/// Hard cap on the number of BIFF records scanned while searching for legacy NOTE comments.
///
/// The `.xls` importer performs multiple best-effort passes over each worksheet substream. Without
/// a cap, a crafted workbook with millions of cell records can force excessive work even when a
/// particular feature (like comments) is absent.
const MAX_RECORDS_SCANNED_PER_SHEET_NOTES_SCAN: usize = 500_000;

/// Maximum number of legacy NOTE/OBJ/TXO note groups to parse per worksheet.
///
/// This bounds memory usage for malicious `.xls` files that contain extremely large numbers of
/// comments.
const MAX_NOTES_PER_SHEET: usize = 20_000;
/// Maximum number of TXO text payloads to store per worksheet while resolving NOTE records.
///
/// In well-formed files this should be <= the number of NOTE records, but we cap it separately to
/// avoid unbounded memory usage when the stream contains many standalone/dangling OBJ/TXO records.
const MAX_TXO_TEXTS_PER_SHEET: usize = 20_000;

const MAX_WARNINGS_PER_SHEET: usize = 20;

fn looks_like_txo_formatting_runs(fragment: &[u8]) -> bool {
    // TXO rich-text formatting runs are stored as an array of 4-byte records:
    // [ich: u16][ifnt: u16] (see MS-XLS 2.5.267 / 2.4.334).
    //
    // When the TXO header is missing/truncated we may not know `cbRuns`, so we use a heuristic
    // to avoid decoding formatting run bytes as text.
    //
    // Importantly: formatting-run CONTINUE payloads do *not* begin with the 1-byte "high-byte"
    // string flag used by continued-string fragments. That means the payload length is usually
    // a multiple of 4, whereas continued-string fragments are typically `1 + n` bytes (and
    // Unicode fragments are always odd-length).
    if fragment.len() < 4 || fragment.len() % 4 != 0 {
        return false;
    }

    let mut likely_records = 0usize;
    let mut total_records = 0usize;
    let mut all_plausible = true;
    let mut is_monotonic = true;
    let mut prev_pos: usize = 0;
    let mut first_pos: Option<usize> = None;
    for chunk in fragment.chunks_exact(4) {
        total_records += 1;
        let pos = u16::from_le_bytes([chunk[0], chunk[1]]) as usize;
        let font = u16::from_le_bytes([chunk[2], chunk[3]]) as usize;

        if first_pos.is_none() {
            first_pos = Some(pos);
            prev_pos = pos;
        } else {
            is_monotonic &= pos >= prev_pos;
            prev_pos = pos;
        }

        if pos > TXO_MAX_TEXT_CHARS || font > 0x0FFF {
            all_plausible = false;
        }

        // Primary heuristic: formatting runs often have zero high bytes and small indices.
        if chunk[1] == 0 && chunk[3] == 0 && pos <= TXO_MAX_TEXT_CHARS && font <= 0x0FFF {
            likely_records += 1;
        }
    }

    // Require a majority match so we don't accidentally treat short/odd continued-string fragments
    // as formatting runs.
    if total_records > 0 && likely_records * 2 >= total_records {
        return true;
    }

    // Secondary heuristic: Some files have many runs with `ich` positions > 255, which means the
    // high byte is non-zero. In that case, the "zero high byte" heuristic above may fail even
    // though the payload is still clearly an array of formatting runs.
    //
    // Be conservative: require
    // - plausible ranges for every run record
    // - monotonic (non-decreasing) `ich` positions
    // - the first run to start at character position 0 (a common invariant for rich-text runs)
    if all_plausible && is_monotonic && first_pos == Some(0) {
        return true;
    }

    false
}

fn split_txo_text_and_formatting_run_suffix(bytes: &[u8]) -> (&[u8], bool) {
    // Some malformed `.xls` writers appear to append TXO rich-text formatting run data directly
    // after the text bytes within the *same* CONTINUE fragment. When the TXO header is truncated
    // (missing `cbRuns`) we don't know where the text ends, and decoding those bytes can leak
    // control characters into the recovered comment text.
    //
    // Best-effort: detect a formatting-run suffix (array of 4-byte records) and strip it.
    //
    // Be conservative:
    // - only consider suffixes whose length is a multiple of 4 bytes
    // - require the first formatting run to start at character position 0 (`ich=0`)
    if bytes.len() < 4 {
        return (bytes, false);
    }

    let max_suffix_len = bytes.len() - (bytes.len() % 4);
    let mut suffix_len = max_suffix_len;
    while suffix_len >= 4 {
        let Some(start) = bytes.len().checked_sub(suffix_len) else {
            break;
        };
        let suffix = &bytes[start..];
        if suffix.len() >= 4
            && suffix[0] == 0
            && suffix[1] == 0
            && looks_like_txo_formatting_runs(suffix)
        {
            return (&bytes[..start], true);
        }
        suffix_len -= 4;
    }

    (bytes, false)
}

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
    parse_biff_sheet_notes_with_record_cap(
        workbook_stream,
        start,
        biff,
        codepage,
        MAX_RECORDS_SCANNED_PER_SHEET_NOTES_SCAN,
    )
}

fn parse_biff_sheet_notes_with_record_cap(
    workbook_stream: &[u8],
    start: usize,
    biff: BiffVersion,
    codepage: u16,
    record_cap: usize,
) -> Result<ParsedSheetNotes, String> {
    let allows_continuation = |record_id: u16| record_id == RECORD_TXO;
    let iter =
        records::LogicalBiffRecordIter::from_offset(workbook_stream, start, allows_continuation)?;

    let mut notes: Vec<ParsedNote> = Vec::new();
    let mut note_obj_ids: HashSet<u16> = HashSet::new();
    let mut texts_by_obj_id: HashMap<u16, String> = HashMap::new();
    let mut current_obj_id: Option<u16> = None;
    let mut warnings: Vec<String> = Vec::new();
    let mut notes_truncated = false;
    let mut texts_truncated = false;
    let mut scanned = 0usize;

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

        scanned = match scanned.checked_add(1) {
            Some(v) => v,
            None => {
                push_warning_force(
                    &mut warnings,
                    "record counter overflow while scanning worksheet notes; stopping early",
                );
                break;
            }
        };
        if scanned > record_cap {
            push_warning_force(
                &mut warnings,
                format!(
                    "too many BIFF records while scanning worksheet notes (cap={record_cap}); stopping early"
                ),
            );
            break;
        }

        match record.record_id {
            RECORD_NOTE => {
                if notes.len() >= MAX_NOTES_PER_SHEET {
                    if !notes_truncated {
                        notes_truncated = true;
                        push_warning(
                            &mut warnings,
                            format!(
                                "too many NOTE records in worksheet; capped at {MAX_NOTES_PER_SHEET}"
                            ),
                        );
                    }
                    // Stop collecting NOTE records to avoid unbounded allocations, but keep scanning
                    // so we can still recover TXO payloads for notes already parsed.
                    continue;
                }

                if let Some(note) = parse_note_record(
                    record.data.as_ref(),
                    record.offset,
                    biff,
                    codepage,
                    &mut warnings,
                ) {
                    note_obj_ids.insert(note.primary_obj_id);
                    note_obj_ids.insert(note.secondary_obj_id);
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

                // If NOTE parsing was truncated, only retain TXO payloads for objects referenced by
                // the notes we kept.
                if notes_truncated && !note_obj_ids.contains(&obj_id) {
                    continue;
                }

                if texts_by_obj_id.len() >= MAX_TXO_TEXTS_PER_SHEET
                    && !texts_by_obj_id.contains_key(&obj_id)
                {
                    if !texts_truncated {
                        texts_truncated = true;
                        push_warning(
                            &mut warnings,
                            format!(
                                "too many TXO text payloads in worksheet; capped at {MAX_TXO_TEXTS_PER_SHEET}"
                            ),
                        );
                    }
                    continue;
                }

                if let Some(text) = parse_txo_text(&record, biff, codepage, &mut warnings) {
                    texts_by_obj_id.insert(obj_id, text);
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    let mut out: Vec<BiffNote> = Vec::new();
    let _ = out.try_reserve_exact(notes.len());
    let mut out_by_obj_id: HashMap<u16, usize> = HashMap::new();
    for note in notes {
        let Some((obj_id, text)) = texts_by_obj_id
            .get(&note.primary_obj_id)
            .map(|text| (note.primary_obj_id, text))
            .or_else(|| {
                texts_by_obj_id
                    .get(&note.secondary_obj_id)
                    .map(|text| (note.secondary_obj_id, text))
            })
        else {
            // No TXO payload for this NOTE record: keep best-effort import going, but skip creating
            // a model comment with missing text.
            push_warning(
                &mut warnings,
                format!(
                    "NOTE record for cell {} references missing TXO payload (obj_id={}, fallback_obj_id={})",
                    {
                        let mut a1 = String::new();
                        formula_model::push_a1_cell_ref(note.cell.row, note.cell.col, false, false, &mut a1);
                        a1
                    },
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
            push_warning(
                &mut warnings,
                format!(
                    "duplicate NOTE record for object id {obj_id} (cell {}); overwriting previous NOTE at cell {}",
                    {
                        let mut a1 = String::new();
                        formula_model::push_a1_cell_ref(resolved.cell.row, resolved.cell.col, false, false, &mut a1);
                        a1
                    },
                    out.get(existing)
                        .map(|note| {
                            let mut a1 = String::new();
                            formula_model::push_a1_cell_ref(note.cell.row, note.cell.col, false, false, &mut a1);
                            a1
                        })
                        .unwrap_or_else(|| "<unknown>".to_string())
                ),
            );
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

fn push_warning_force(warnings: &mut Vec<String>, warning: impl Into<String>) {
    let warning = warning.into();
    if warnings.len() < MAX_WARNINGS_PER_SHEET {
        warnings.push(warning);
    } else if let Some(last) = warnings.last_mut() {
        *last = warning;
    }
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
            // Most files match the spec-defined string encoding:
            // - BIFF8: ShortXLUnicodeString
            // - BIFF5: ANSI short string
            //
            // But files in the wild sometimes store NOTE authors using BIFF8 string encodings even
            // in BIFF5 workbooks (e.g. a leading flags byte), or store an `XLUnicodeString`
            // (16-bit length). If the spec-defined parser doesn't consume the full payload, try
            // those alternatives best-effort.
            if consumed != author_bytes.len() {
                match biff {
                    BiffVersion::Biff8 => match strings::parse_biff8_unicode_string(author_bytes, codepage) {
                        Ok((alt, alt_consumed)) if alt_consumed == author_bytes.len() => alt,
                        _ => s,
                    },
                    BiffVersion::Biff5 => {
                        // Some BIFF5 writers store NOTE authors using BIFF8 string encodings.
                        // Attempt both common BIFF8 variants (ShortXLUnicodeString / XLUnicodeString)
                        // when the BIFF5 short-string parser doesn't consume the full payload.
                        let mut recovered: Option<String> = None;
                        if let Ok((alt, alt_consumed)) =
                            strings::parse_biff8_short_string(author_bytes, codepage)
                        {
                            if alt_consumed == author_bytes.len() {
                                recovered = Some(alt);
                            }
                        }
                        if recovered.is_none() {
                            if let Ok((alt, alt_consumed)) =
                                strings::parse_biff8_unicode_string(author_bytes, codepage)
                            {
                                if alt_consumed == author_bytes.len() {
                                    recovered = Some(alt);
                                }
                            }
                        }

                        recovered.unwrap_or(s)
                    }
                }
            } else {
                s
            }
        }
        Err(err) => match biff {
            BiffVersion::Biff8 => match strings::parse_biff8_unicode_string(author_bytes, codepage) {
                Ok((alt, _)) => alt,
                Err(unicode_err) => {
                    if let Some(best_effort) =
                        strings::parse_biff5_short_string_best_effort(author_bytes, codepage)
                    {
                        push_warning(
                            warnings,
                            format!(
                                "failed to parse NOTE author string at offset {offset}: {err}; XLUnicodeString fallback also failed: {unicode_err}; treating author as BIFF5 ANSI short string"
                            ),
                        );
                        best_effort
                    } else {
                        push_warning(
                            warnings,
                            format!(
                                "failed to parse NOTE author string at offset {offset}: {err}; XLUnicodeString fallback also failed: {unicode_err}"
                            ),
                        );
                        String::new()
                    }
                }
            },
            BiffVersion::Biff5 => {
                // Best-effort: some BIFF5 files appear to store NOTE authors using BIFF8 string
                // encodings (leading option flags byte, 16-bit length prefix, etc).
                if let Ok((alt, alt_consumed)) =
                    strings::parse_biff8_short_string(author_bytes, codepage)
                {
                    if alt_consumed == author_bytes.len() {
                        alt
                    } else if let Ok((alt, alt_consumed)) =
                        strings::parse_biff8_unicode_string(author_bytes, codepage)
                    {
                        if alt_consumed == author_bytes.len() {
                            alt
                        } else if let Some(best_effort) =
                            strings::parse_biff5_short_string_best_effort(author_bytes, codepage)
                        {
                            push_warning(
                                warnings,
                                format!(
                                    "failed to fully parse NOTE author string at offset {offset}: {err}; treating author as truncated BIFF5 ANSI short string"
                                ),
                            );
                            best_effort
                        } else {
                            push_warning(
                                warnings,
                                format!(
                                    "failed to fully parse NOTE author string at offset {offset}: {err}"
                                ),
                            );
                            String::new()
                        }
                    } else if let Some(best_effort) =
                        strings::parse_biff5_short_string_best_effort(author_bytes, codepage)
                    {
                        push_warning(
                            warnings,
                            format!(
                                "failed to fully parse NOTE author string at offset {offset}: {err}; treating author as truncated BIFF5 ANSI short string"
                            ),
                        );
                        best_effort
                    } else {
                        push_warning(
                            warnings,
                            format!(
                                "failed to fully parse NOTE author string at offset {offset}: {err}"
                            ),
                        );
                        String::new()
                    }
                } else if let Ok((alt, alt_consumed)) =
                    strings::parse_biff8_unicode_string(author_bytes, codepage)
                {
                    if alt_consumed == author_bytes.len() {
                        alt
                    } else if let Some(best_effort) =
                        strings::parse_biff5_short_string_best_effort(author_bytes, codepage)
                    {
                        push_warning(
                            warnings,
                            format!(
                                "failed to fully parse NOTE author string at offset {offset}: {err}; treating author as truncated BIFF5 ANSI short string"
                            ),
                        );
                        best_effort
                    } else {
                        push_warning(
                            warnings,
                            format!(
                                "failed to fully parse NOTE author string at offset {offset}: {err}"
                            ),
                        );
                        String::new()
                    }
                } else if let Some(best_effort) =
                    strings::parse_biff5_short_string_best_effort(author_bytes, codepage)
                {
                    push_warning(
                        warnings,
                        format!(
                            "failed to parse NOTE author string at offset {offset}: {err}; treating author as truncated BIFF5 ANSI short string"
                        ),
                    );
                    best_effort
                } else {
                    push_warning(
                        warnings,
                        format!("failed to parse NOTE author string at offset {offset}: {err}"),
                    );
                    String::new()
                }
            }
        },
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

    while let Some(header) = data.get(idx..).and_then(|rest| rest.get(..4)) {
        let ft = u16::from_le_bytes([header[0], header[1]]);
        let cb = u16::from_le_bytes([header[2], header[3]]) as usize;
        idx = match idx.checked_add(4) {
            Some(v) => v,
            None => {
                push_warning(
                    warnings,
                    format!("OBJ record at offset {record_offset} has subrecord offset overflow"),
                );
                break;
            }
        };

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
                format!(
                    "OBJ record at offset {record_offset} has truncated subrecord 0x{ft:04X} (cb={cb})"
                ),
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
    let first = record.first_fragment();
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let continues = fragments.get(1..).unwrap_or_default();
    if continues.is_empty() {
        match parse_txo_cch_text_biff5(first, 0) {
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

    // BIFF5 typically stores the TXO text bytes directly in subsequent CONTINUE records (no
    // per-fragment option flags byte). Treat the continued bytes as ANSI encoded using the
    // workbook codepage.
    //
    // Some producers appear to mimic BIFF8's continued-string layout and prefix each CONTINUE
    // fragment with a one-byte "high-byte" flag (0/1). In that case, the TXO `cchText` count does
    // *not* include those flag bytes, so treat them as optional and skip them best-effort.

    // Like BIFF8, some BIFF5 files reserve trailing continuation bytes for rich-text formatting
    // runs (per TXO `cbRuns`). Respect that so we don't decode formatting-run bytes as text when
    // `cchText` is larger than the available text bytes.
    let cb_runs = TXO_RUNS_LEN_OFFSET
        .checked_add(2)
        .and_then(|end| first.get(TXO_RUNS_LEN_OFFSET..end))
        .map(|v| u16::from_le_bytes([v[0], v[1]]) as usize);
    let has_cb_runs = cb_runs.is_some();
    let cb_runs = cb_runs.unwrap_or(0);
    let total_continue_bytes: usize = continues.iter().map(|frag| frag.len()).sum();
    let mut text_continue_bytes = total_continue_bytes.checked_sub(cb_runs).unwrap_or(0);
    if cb_runs > total_continue_bytes {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} has cbRuns ({cb_runs}) larger than continuation payload ({total_continue_bytes})",
                record.offset
            ),
        );
        // Best-effort: if the header's `cbRuns` is corrupt and exceeds the observed continuation
        // payload, ignore it and fall back to heuristics (`looks_like_txo_formatting_runs`) to stop
        // before formatting-run bytes.
        text_continue_bytes = total_continue_bytes;
    }

    // Compute capacities using only the text-bytes region (excluding the trailing cbRuns bytes).
    let mut capacity_raw = 0usize;
    let mut remaining_bytes = text_continue_bytes;
    for &frag in continues {
        if remaining_bytes == 0 {
            break;
        }
        let take_len = frag.len().min(remaining_bytes);
        remaining_bytes -= take_len;
        let frag = &frag[..take_len];
        if looks_like_txo_formatting_runs(frag) {
            break;
        }
        capacity_raw = match capacity_raw.checked_add(take_len) {
            Some(v) => v.min(TXO_MAX_TEXT_CHARS),
            None => TXO_MAX_TEXT_CHARS,
        };
    }

    let first_continue_has_flag = matches!(
        continues
            .first()
            .copied()
            .and_then(|frag| frag.first().copied()),
        Some(0) | Some(1)
    );
    let capacity_without_flags = if first_continue_has_flag {
        let mut cap = 0usize;
        let mut remaining_bytes = text_continue_bytes;
        for &frag in continues {
            if remaining_bytes == 0 {
                break;
            }
            let take_len = frag.len().min(remaining_bytes);
            remaining_bytes -= take_len;
            let frag = &frag[..take_len];
            if looks_like_txo_formatting_runs(frag) {
                break;
            }
            let len = if matches!(frag.first().copied(), Some(0) | Some(1)) {
                frag.len().checked_sub(1).unwrap_or(0)
            } else {
                frag.len()
            };
            cap = match cap.checked_add(len) {
                Some(v) => v.min(TXO_MAX_TEXT_CHARS),
                None => TXO_MAX_TEXT_CHARS,
            };
        }
        cap
    } else {
        capacity_raw
    };

    let spec_cch_text = TXO_TEXT_LEN_OFFSET
        .checked_add(2)
        .and_then(|end| first.get(TXO_TEXT_LEN_OFFSET..end))
        .map(|v| u16::from_le_bytes([v[0], v[1]]) as usize)
        .filter(|&cch| cch != 0 && cch <= TXO_MAX_TEXT_CHARS);
    // Prefer the spec-defined cchText field (offset 6) when present, even if it exceeds the
    // observed continuation capacity (truncated/corrupt files). This preserves truncation
    // warnings and avoids accidentally selecting another header field that happens to fit the
    // available payload.
    let cch_text = spec_cch_text.or_else(|| detect_txo_cch_text(first, capacity_raw));

    let skip_leading_flag_bytes = match cch_text {
        Some(cch) => first_continue_has_flag && cch <= capacity_without_flags,
        None => first_continue_has_flag,
    };
    let capacity = if skip_leading_flag_bytes {
        capacity_without_flags
    } else {
        capacity_raw
    };

    let mut remaining = match cch_text {
        Some(cch) => cch,
        None => {
            if capacity > 0 {
                push_warning(
                    warnings,
                    format!(
                        "TXO record at offset {} has malformed header/cchText; falling back to decoding CONTINUE fragments",
                        record.offset
                    ),
                );
            }
            capacity
        }
    };
    if remaining == 0 {
        return Some(String::new());
    }

    // Accumulate the byte payload first, then decode once. This preserves stateful multibyte
    // codepages (e.g. Shift-JIS) when a character boundary is split across CONTINUE records.
    let mut bytes = Vec::new();
    let _ = bytes.try_reserve_exact(remaining);
    let mut remaining_bytes = text_continue_bytes;
    for &frag in continues {
        if remaining == 0 || remaining_bytes == 0 {
            break;
        }
        if frag.is_empty() {
            continue;
        }

        let take_len = frag.len().min(remaining_bytes);
        remaining_bytes -= take_len;
        let frag = &frag[..take_len];
        if looks_like_txo_formatting_runs(frag) {
            break;
        }
        let mut frag = frag;

        if skip_leading_flag_bytes && matches!(frag.first().copied(), Some(0) | Some(1)) {
            frag = frag.get(1..).unwrap_or(&[]);
        }
        if frag.is_empty() {
            continue;
        }

        // If `cbRuns` is missing (truncated header), some malformed files still append formatting
        // run bytes to the end of the *text* fragment. Strip such suffixes best-effort so we don't
        // decode them as characters.
        let mut frag = frag;
        let mut take = remaining.min(frag.len());
        if !has_cb_runs && take == frag.len() {
            let (trimmed, _) = split_txo_text_and_formatting_run_suffix(frag);
            frag = trimmed;
            take = remaining.min(frag.len());
        }
        if take == 0 {
            continue;
        }
        bytes.extend_from_slice(&frag[..take]);
        remaining -= take;
    }

    if let Some(cch) = cch_text {
        if remaining > 0 {
            push_warning(
                warnings,
                format!(
                    "TXO record at offset {} truncated text (wanted {cch} chars, got {})",
                    record.offset,
                    cch.checked_sub(remaining).unwrap_or(0)
                ),
            );
        }
    }

    let mut out = strings::decode_ansi(codepage, &bytes);
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

    // Some nonstandard `.xls` producers appear to store BIFF8 TXO text continuation bytes without
    // the leading "high-byte" option flags byte (i.e. using BIFF5-style raw bytes). If the first
    // continuation fragment doesn't start with a plausible flags value, fall back to the BIFF5
    // decoder which treats continuation bytes as ANSI/codepage bytes.
    if let Some(first_continue_flags) = continues.first().and_then(|frag| frag.first().copied()) {
        if !matches!(first_continue_flags, 0 | 1) {
            push_warning(
                warnings,
                format!(
                    "TXO record at offset {} has nonstandard CONTINUE text payload (missing BIFF8 flags byte); decoding as BIFF5 ANSI text",
                    record.offset
                ),
            );
            return parse_txo_text_biff5(record, codepage, warnings);
        }
    }

    // `cbRuns` indicates how many bytes at the end of the TXO continuation area are reserved for
    // rich-text formatting runs. We ignore those bytes so we don't misinterpret formatting run data
    // as characters if `cchText` is larger than the available text bytes (truncated/corrupt files).
    let cb_runs = TXO_RUNS_LEN_OFFSET
        .checked_add(2)
        .and_then(|end| first.get(TXO_RUNS_LEN_OFFSET..end))
        .map(|v| u16::from_le_bytes([v[0], v[1]]) as usize);
    let has_cb_runs = cb_runs.is_some();
    let cb_runs = cb_runs.unwrap_or(0);
    let total_continue_bytes: usize = continues.iter().map(|frag| frag.len()).sum();
    let mut text_continue_bytes = total_continue_bytes.checked_sub(cb_runs).unwrap_or(0);
    if cb_runs > total_continue_bytes {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} has cbRuns ({cb_runs}) larger than continuation payload ({total_continue_bytes})",
                record.offset
            ),
        );
        // Best-effort: if the header's `cbRuns` is corrupt and exceeds the observed continuation
        // payload, ignore it and fall back to heuristics (`looks_like_txo_formatting_runs`) to stop
        // before formatting-run bytes.
        text_continue_bytes = total_continue_bytes;
    }

    let max_chars = estimate_max_chars_with_byte_limit(continues, text_continue_bytes);
    let Some(cch_text) = parse_txo_cch_text(first, max_chars) else {
        return fallback_decode_continue_fragments(record, codepage, warnings);
    };
    if cch_text == 0 {
        return Some(String::new());
    }

    let mut out = String::new();
    let mut ansi_bytes: Vec<u8> = Vec::new();
    let mut utf16_bytes: Vec<u8> = Vec::new();
    let mut pending_unicode_byte: Option<u8> = None;
    let mut current_is_unicode: Option<bool> = None;
    let mut warned_missing_flags = false;
    let mut remaining = cch_text;
    let mut remaining_bytes = text_continue_bytes;
    for frag in continues {
        if remaining == 0 || remaining_bytes == 0 {
            break;
        }

        let take_len = frag.len().min(remaining_bytes);
        remaining_bytes -= take_len;
        let frag = &frag[..take_len];
        if looks_like_txo_formatting_runs(frag) {
            break;
        }
        let Some((&first, rest)) = frag.split_first() else {
            continue;
        };
        let (mut bytes, is_unicode, has_flags) = if matches!(first, 0 | 1) {
            let is_unicode = (first & 0x01) != 0;
            (rest, is_unicode, true)
        } else {
            // Nonstandard fragment: missing the 1-byte "high-byte" flag. Assume this fragment
            // continues using the same encoding as the previous fragment (default: compressed).
            let is_unicode = current_is_unicode.unwrap_or(false);
            (frag, is_unicode, false)
        };
        if has_flags {
            current_is_unicode = Some(is_unicode);
        } else if !warned_missing_flags {
            warned_missing_flags = true;
            push_warning(
                warnings,
                format!(
                    "TXO record at offset {} has CONTINUE fragment missing BIFF8 flags byte; assuming the previous fragment encoding",
                    record.offset
                ),
            );
        }
        if is_unicode {
            // Flush any buffered ANSI bytes so multi-byte sequences can span CONTINUE fragments.
            if !ansi_bytes.is_empty() {
                out.push_str(&strings::decode_ansi(codepage, &ansi_bytes));
                ansi_bytes.clear();
            }

            if !has_cb_runs {
                // If `cbRuns` is missing, some malformed files append formatting run bytes directly
                // after UTF-16LE text bytes within the same fragment. When we would otherwise
                // consume the full fragment (because `cchText` is too large), strip a formatting-run
                // suffix best-effort so we don't decode it as text.
                let combined_len = bytes.len() + usize::from(pending_unicode_byte.is_some());
                let available_chars = combined_len / 2;
                if available_chars > 0 && remaining >= available_chars {
                    let (trimmed, _) = split_txo_text_and_formatting_run_suffix(bytes);
                    bytes = trimmed;
                }
            }

            // Best-effort: some files split UTF-16LE code units across CONTINUE fragments. Buffer
            // a single leftover byte so we can recover the character when the next fragment arrives.
            let combined_len = bytes.len() + usize::from(pending_unicode_byte.is_some());
            let available_chars = combined_len / 2;
            if available_chars == 0 {
                if pending_unicode_byte.is_none() {
                    if let Some(&b) = bytes.first() {
                        pending_unicode_byte = Some(b);
                    }
                }
                continue;
            }

            let take_chars = remaining.min(available_chars);
            let take_bytes_total = take_chars * 2;
            let mut buf = Vec::new();
            let _ = buf.try_reserve_exact(take_bytes_total);
            if let Some(b) = pending_unicode_byte.take() {
                buf.push(b);
            }
            let need_from_current = take_bytes_total.checked_sub(buf.len()).unwrap_or(0);
            let take_from_current = need_from_current.min(bytes.len());
            buf.extend_from_slice(&bytes[..take_from_current]);
            let used_current = take_from_current;
            // Buffer all UTF-16LE bytes across fragments so surrogate pairs split across CONTINUE
            // records can still be decoded correctly.
            utf16_bytes.extend_from_slice(&buf);

            remaining -= take_chars;

            // Preserve any trailing odd byte for the next UTF-16LE fragment (only relevant when
            // we haven't satisfied `cchText` yet).
            if remaining > 0 && bytes.len() > used_current {
                pending_unicode_byte = bytes.get(used_current).copied();
            } else {
                pending_unicode_byte = None;
            }
        } else {
            if !utf16_bytes.is_empty() {
                out.push_str(&decode_utf16le(&utf16_bytes));
                utf16_bytes.clear();
            }
            pending_unicode_byte = None;
            // Accumulate and decode once to preserve stateful multibyte encodings (e.g. Shift-JIS)
            // when a character boundary is split across CONTINUE records.
            if !has_cb_runs && remaining >= bytes.len() {
                let (trimmed, _) = split_txo_text_and_formatting_run_suffix(bytes);
                bytes = trimmed;
            }
            let available_chars = bytes.len();
            if available_chars == 0 {
                continue;
            }
            let take_chars = remaining.min(available_chars);
            let slice = &bytes[..take_chars];
            ansi_bytes.extend_from_slice(slice);
            remaining -= take_chars;
        }
    }

    if remaining > 0 {
        push_warning(
            warnings,
            format!(
                "TXO record at offset {} truncated text (wanted {cch_text} chars, got {})",
                record.offset,
                cch_text.checked_sub(remaining).unwrap_or(0)
            ),
        );
    }
    if !utf16_bytes.is_empty() {
        out.push_str(&decode_utf16le(&utf16_bytes));
    }
    if !ansi_bytes.is_empty() {
        out.push_str(&strings::decode_ansi(codepage, &ansi_bytes));
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

fn detect_txo_cch_text(header: &[u8], continue_capacity: usize) -> Option<usize> {
    if continue_capacity == 0 {
        return None;
    }

    for off in TXO_TEXT_LEN_OFFSETS {
        let Some(bytes) = off
            .checked_add(2)
            .and_then(|end| header.get(off..end))
        else {
            continue;
        };
        let cch = u16::from_le_bytes([bytes[0], bytes[1]]) as usize;
        if cch == 0 {
            continue;
        }
        if cch > TXO_MAX_TEXT_CHARS {
            continue;
        }
        if cch <= continue_capacity {
            return Some(cch);
        }
    }

    None
}

fn parse_txo_cch_text(header: &[u8], max_chars: usize) -> Option<usize> {
    // Heuristic: cchText is typically at offset 6 in the TXO header, but some sources disagree.
    // Try a few common offsets and choose the first plausible value.
    if header.len() < 8 {
        return None;
    }

    let max_chars = max_chars.min(TXO_MAX_TEXT_CHARS);

    // Spec-defined BIFF8 offset for cchText.
    let mut cch_at_6 = TXO_TEXT_LEN_OFFSET
        .checked_add(2)
        .and_then(|end| header.get(TXO_TEXT_LEN_OFFSET..end))
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) as usize)
        .unwrap_or(0);
    if cch_at_6 > TXO_MAX_TEXT_CHARS {
        cch_at_6 = 0;
    }

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
    //
    // Note: we intentionally *exclude* offset 12 because that's the spec-defined `cbRuns` field.
    // Some files report `cchText=0` while still setting `cbRuns=4`; treating `cbRuns` as an
    // alternate text length would incorrectly truncate the recovered comment text.
    for off in [4usize, 8usize, 10usize] {
        let Some(bytes) = off
            .checked_add(2)
            .and_then(|end| header.get(off..end))
        else {
            continue;
        };
        let cch = u16::from_le_bytes([bytes[0], bytes[1]]) as usize;
        if cch != 0 && cch <= max_chars && cch <= TXO_MAX_TEXT_CHARS {
            return Some(cch);
        }
    }

    // Last resort: decode as much as we can from the available continuation bytes.
    Some(max_chars)
}

fn parse_txo_cch_text_biff5(header: &[u8], max_chars: usize) -> Option<usize> {
    // BIFF5 uses the same TXO `cchText` field conceptually, but the CONTINUE payload is typically
    // raw 8-bit text bytes (no per-fragment option flags). We can reuse the BIFF8 heuristic logic
    // as long as `max_chars` is computed as a simple byte count.
    parse_txo_cch_text(header, max_chars)
}

fn estimate_max_chars_with_byte_limit(continues: &[&[u8]], byte_limit: usize) -> usize {
    // Best-effort estimate used only for header heuristics.
    //
    // Most BIFF8 TXO continuation fragments begin with a 1-byte "high-byte" flag (0/1) followed
    // by the text bytes. Some malformed files omit that flag in one or more fragments, and some
    // files split UTF-16LE code units across fragment boundaries. Keep this estimate robust enough
    // that we don't under-count the available text bytes (which would silently truncate comment
    // text when `cchText` is missing/zero).
    let mut total = 0usize;
    let mut remaining = byte_limit;
    let mut current_is_unicode: Option<bool> = None;
    let mut pending_unicode_byte = false;
    for frag in continues {
        if remaining == 0 {
            break;
        }
        let take_len = frag.len().min(remaining);
        remaining -= take_len;
        let frag = &frag[..take_len];
        if frag.is_empty() {
            continue;
        };

        if looks_like_txo_formatting_runs(frag) {
            break;
        }

        let first = frag[0];
        let (bytes, is_unicode) = if matches!(first, 0 | 1) {
            let is_unicode = (first & 0x01) != 0;
            current_is_unicode = Some(is_unicode);
            (&frag[1..], is_unicode)
        } else {
            // Missing flags byte: assume the previous fragment encoding (default: compressed).
            let is_unicode = current_is_unicode.unwrap_or(false);
            (frag, is_unicode)
        };

        if is_unicode {
            let combined_len = bytes.len() + usize::from(pending_unicode_byte);
            total = total
                .checked_add(combined_len / 2)
                .unwrap_or(usize::MAX)
                .min(TXO_MAX_TEXT_CHARS);
            pending_unicode_byte = combined_len % 2 == 1;
        } else {
            pending_unicode_byte = false;
            total = total
                .checked_add(bytes.len())
                .unwrap_or(usize::MAX)
                .min(TXO_MAX_TEXT_CHARS);
        }
    }
    total
}

fn fallback_decode_continue_fragments(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
    warnings: &mut Vec<String>,
) -> Option<String> {
    let mut fragments = record.fragments();
    let _ = fragments.next(); // skip header

    let continues: Vec<&[u8]> = fragments.collect();
    if continues.is_empty() {
        push_warning(
            warnings,
            format!("TXO record at offset {} missing CONTINUE fragments", record.offset),
        );
        return Some(String::new());
    };

    push_warning(
        warnings,
        format!(
            "TXO record at offset {} has malformed header; falling back to decoding CONTINUE fragments",
            record.offset
        ),
    );

    let mut out = String::new();
    let mut ansi_bytes: Vec<u8> = Vec::new();
    let mut utf16_bytes: Vec<u8> = Vec::new();
    let mut pending_unicode_byte: Option<u8> = None;
    let mut current_is_unicode: Option<bool> = None;
    let mut warned_missing_flags = false;
    let mut remaining_chars = TXO_MAX_TEXT_CHARS;
    for frag in continues {
        if remaining_chars == 0 {
            break;
        }

        if frag.is_empty() {
            continue;
        }

        // If this fragment looks like TXO formatting run data (no leading flags byte), stop before
        // we accidentally decode it as text.
        if looks_like_txo_formatting_runs(frag) {
            break;
        }

        let Some((&first, rest)) = frag.split_first() else {
            continue;
        };

        let (bytes, is_unicode, has_flags) = if matches!(first, 0 | 1) {
            let is_unicode = (first & 0x01) != 0;
            (rest, is_unicode, true)
        } else {
            // Nonstandard fragment: missing the 1-byte "high-byte" flag. Assume this fragment
            // continues using the same encoding as the previous fragment (default: compressed).
            let is_unicode = current_is_unicode.unwrap_or(false);
            (frag, is_unicode, false)
        };

        if has_flags {
            current_is_unicode = Some(is_unicode);
        } else if !warned_missing_flags {
            warned_missing_flags = true;
            push_warning(
                warnings,
                format!(
                    "TXO record at offset {} has CONTINUE fragment missing BIFF8 flags byte; assuming the previous fragment encoding",
                    record.offset
                ),
            );
        }

        let mut bytes = bytes;
        if is_unicode {
            if !ansi_bytes.is_empty() {
                out.push_str(&strings::decode_ansi(codepage, &ansi_bytes));
                ansi_bytes.clear();
            }

            let combined_len = bytes.len() + usize::from(pending_unicode_byte.is_some());
            let available_chars = combined_len / 2;
            if available_chars > 0 && remaining_chars >= available_chars {
                let (trimmed, _) = split_txo_text_and_formatting_run_suffix(bytes);
                bytes = trimmed;
            }

            let combined_len = bytes.len() + usize::from(pending_unicode_byte.is_some());
            let available_chars = combined_len / 2;
            if available_chars == 0 {
                if pending_unicode_byte.is_none() {
                    if let Some(&b) = bytes.first() {
                        pending_unicode_byte = Some(b);
                    }
                }
                continue;
            }

            let take_chars = remaining_chars.min(available_chars);
            let take_bytes_total = take_chars * 2;
            let mut buf = Vec::new();
            let _ = buf.try_reserve_exact(take_bytes_total);
            if let Some(b) = pending_unicode_byte.take() {
                buf.push(b);
            }
            let need_from_current = take_bytes_total.checked_sub(buf.len()).unwrap_or(0);
            let take_from_current = need_from_current.min(bytes.len());
            buf.extend_from_slice(&bytes[..take_from_current]);
            let used_current = take_from_current;
            // Buffer all UTF-16LE bytes across fragments so surrogate pairs split across CONTINUE
            // records can still be decoded correctly.
            utf16_bytes.extend_from_slice(&buf);
            remaining_chars -= take_chars;

            if remaining_chars > 0 && bytes.len() > used_current {
                pending_unicode_byte = bytes.get(used_current).copied();
            } else {
                pending_unicode_byte = None;
            }
        } else {
            if !utf16_bytes.is_empty() {
                out.push_str(&decode_utf16le(&utf16_bytes));
                utf16_bytes.clear();
            }
            pending_unicode_byte = None;
            if remaining_chars >= bytes.len() {
                let (trimmed, _) = split_txo_text_and_formatting_run_suffix(bytes);
                bytes = trimmed;
            }
            let available_chars = bytes.len();
            if available_chars == 0 {
                continue;
            }

            let take_chars = remaining_chars.min(available_chars);
            let slice = &bytes[..take_chars];
            ansi_bytes.extend_from_slice(slice);
            remaining_chars -= take_chars;
        }
    }
    if !utf16_bytes.is_empty() {
        out.push_str(&decode_utf16le(&utf16_bytes));
    }
    if !ansi_bytes.is_empty() {
        out.push_str(&strings::decode_ansi(codepage, &ansi_bytes));
    }

    trim_trailing_nuls(&mut out);
    strip_embedded_nuls(&mut out);
    Some(out)
}

fn decode_utf16le(bytes: &[u8]) -> String {
    let mut u16s = Vec::new();
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
        let mut out = Vec::new();
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
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&cch_text.to_le_bytes());
        record(RECORD_TXO, &payload)
    }

    fn txo_with_cch_text_and_cb_runs(cch_text: u16, cb_runs: u16) -> Vec<u8> {
        let mut payload = vec![0u8; 18];
        payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&cch_text.to_le_bytes());
        payload[TXO_RUNS_LEN_OFFSET..TXO_RUNS_LEN_OFFSET + 2]
            .copy_from_slice(&cb_runs.to_le_bytes());
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
    fn note_record_parses_author_as_biff5_short_string_when_biff8_parsing_fails() {
        // Some nonstandard `.xls` producers appear to store the NOTE author as a BIFF5-style
        // short ANSI string (length + bytes) even in a BIFF8 worksheet. Best-effort parsing should
        // still recover the author.
        let author = "Alice";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // row
        payload.extend_from_slice(&0u16.to_le_bytes()); // col
        payload.extend_from_slice(&1u16.to_le_bytes()); // grbit/idObj
        payload.extend_from_slice(&1u16.to_le_bytes()); // idObj/grbit

        // BIFF5-style author string (no BIFF8 flags byte).
        payload.push(author.len() as u8);
        payload.extend_from_slice(author.as_bytes());

        let mut warnings = Vec::new();
        let parsed = parse_note_record(&payload, 0, BiffVersion::Biff8, 1252, &mut warnings)
            .expect("parse note");
        assert_eq!(parsed.author, author);
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("treating author as BIFF5 ANSI short string")),
            "expected BIFF5 fallback warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn note_record_biff5_parses_author_encoded_as_biff8_short_string() {
        // Some BIFF5 producers appear to store NOTE authors using the BIFF8 ShortXLUnicodeString
        // encoding (length + option flags byte). Ensure we can recover the full author string.
        let author = "Alice";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // row
        payload.extend_from_slice(&0u16.to_le_bytes()); // col
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&1u16.to_le_bytes()); // idObj

        // BIFF8 ShortXLUnicodeString author (compressed) embedded in a BIFF5 NOTE record.
        payload.push(author.len() as u8);
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(author.as_bytes());

        let mut warnings = Vec::new();
        let parsed = parse_note_record(&payload, 0, BiffVersion::Biff5, 1252, &mut warnings)
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        let note = &notes[0];
        assert_eq!(note.cell, CellRef::new(0, 0));
        assert_eq!(note.obj_id, 1);
        assert_eq!(note.author, "Alice");
        assert_eq!(note.text, "Hello");
    }

    #[test]
    fn parses_biff8_txo_text_when_continue_payload_is_missing_flags_byte() {
        // Some nonstandard `.xls` producers appear to store the TXO continuation bytes without
        // the leading BIFF8 "high-byte" flags byte. Best-effort parsing should fall back to
        // treating the payload as BIFF5-style ANSI bytes.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(5, 4),
            // CONTINUE payload contains the text bytes directly (no flags byte).
            record(records::RECORD_CONTINUE, b"Hello"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("missing BIFF8 flags byte") || w.contains("decoding as BIFF5")),
            "expected missing-flags warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_biff8_txo_text_when_a_later_continue_fragment_is_missing_flags_byte() {
        // Similar to `parses_biff8_txo_text_when_continue_payload_is_missing_flags_byte`, but
        // only the *second* fragment omits the flags byte. We should keep decoding and treat the
        // fragment as using the previous encoding (compressed ANSI in this fixture).
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(5, 4),
            continue_text_compressed_bytes(b"He"),
            // Second fragment omits the flags byte.
            record(records::RECORD_CONTINUE, b"llo"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("missing BIFF8 flags byte")),
            "expected missing-flags warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_biff8_txo_text_when_cchtext_is_zero_and_a_later_continue_fragment_is_missing_flags_byte(
    ) {
        // When the TXO header reports `cchText=0`, we infer the length from the continuation area.
        // Ensure the inference still accounts for the full text bytes even if a later fragment is
        // missing the BIFF8 flags byte.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(0, 4),
            continue_text_compressed_bytes(b"He"),
            // Second fragment omits the flags byte.
            record(records::RECORD_CONTINUE, b"llo"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("missing BIFF8 flags byte")),
            "expected missing-flags warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_biff8_txo_text_when_cb_runs_exceeds_continuation_payload() {
        // Some corrupt files contain a `cbRuns` value that exceeds the actual continuation payload.
        // Best-effort parsing should ignore the invalid `cbRuns` (rather than treating all bytes as
        // formatting runs) so we can still recover the text.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(5, 1000),
            continue_text_compressed_bytes(b"Hello"),
            // Formatting runs CONTINUE payload (dummy bytes).
            record(records::RECORD_CONTINUE, &[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("cbRuns (1000)")),
            "expected cbRuns warning; warnings={warnings:?}"
        );
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hi А");
    }

    #[test]
    fn parses_biff5_txo_text_when_cb_runs_exceeds_continuation_payload() {
        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(5, 1000),
            continue_text_biff5(b"Hello"),
            // Formatting CONTINUE payload (dummy bytes).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("cbRuns (1000)")),
            "expected cbRuns warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn does_not_decode_biff5_formatting_run_bytes_as_text_when_cbruns_is_set() {
        // `cchText` is intentionally larger than the available text bytes. Without respecting
        // `cbRuns`, parsers may decode the formatting run bytes as if they were characters.
        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            // cchText=6 but only 5 chars of actual text bytes.
            txo_with_cch_text_and_cb_runs(6, 4),
            continue_text_biff5(b"Hello"),
            // Formatting CONTINUE payload (dummy bytes).
            continue_text_biff5(&[0xFFu8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn biff5_prefers_spec_cchtext_even_when_another_offset_fits_capacity() {
        // Some sources disagree on the TXO header layout and place `cchText` at offset 4.
        //
        // If the spec-defined `cchText` at offset 6 is present but larger than the continuation
        // payload (truncated file), we still want to trust it so we:
        // - decode all available bytes, and
        // - emit a truncation warning.
        //
        // If we instead pick the smaller offset-4 field just because it fits the observed payload,
        // we'd silently drop bytes and miss the truncation warning.
        let mut txo_payload = vec![0u8; 18];
        // Spec-defined field says 5 chars (but the continuation will only contain 3).
        txo_payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&5u16.to_le_bytes());
        // Alternate offset 4 contains a smaller plausible value.
        txo_payload[4..6].copy_from_slice(&2u16.to_le_bytes());

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_payload),
            continue_text_biff5(b"ABC"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "ABC");
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn does_not_decode_biff5_formatting_runs_as_text_when_txo_header_is_truncated() {
        // When the TXO header is truncated such that `cbRuns` is missing, we still want to avoid
        // decoding formatting run bytes as part of the text payload.
        let mut txo_header = vec![0u8; 8];
        txo_header[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&10u16.to_le_bytes());

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_header),
            continue_text_biff5(b"Hello"),
            // Formatting run bytes: [ich=0][ifnt=1].
            continue_text_biff5(&[0x00, 0x00, 0x01, 0x00]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn does_not_decode_biff5_formatting_runs_appended_to_text_fragment_when_cb_runs_is_missing() {
        // Similar to `does_not_decode_biff5_formatting_runs_as_text_when_txo_header_is_truncated`,
        // but the formatting run bytes are appended to the end of the text bytes within the same
        // CONTINUE fragment.
        let mut txo_header = vec![0u8; 8];
        txo_header[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&10u16.to_le_bytes());

        let mut mixed = Vec::new();
        mixed.extend_from_slice(b"Hello");
        // Formatting run bytes: [ich=0][ifnt=1].
        mixed.extend_from_slice(&[0x00, 0x00, 0x01, 0x00]);

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_header),
            continue_text_biff5(&mixed),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn parses_biff5_txo_text_when_cchtext_is_stored_at_alternate_offset() {
        // The spec-defined cchText field (offset 6) is zero, but offset 4 contains the correct
        // value. Best-effort decoding should still recover the full text.
        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_at_offset_4(5),
            continue_text_biff5(b"Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn parses_biff5_txo_text_when_cchtext_is_zero_but_continue_has_text() {
        // Some files report `cchText=0` in the TXO header even though the continuation area still
        // contains the text. Best-effort decoding should still recover it.
        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(0),
            continue_text_biff5(b"Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
        assert!(
            warnings.iter().any(|w| w.contains("malformed header/cchText")),
            "expected fallback warning; warnings={warnings:?}"
        );
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
    fn falls_back_to_decoding_continue_fragments_when_txo_header_is_missing() {
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
    fn falls_back_to_decoding_continue_fragments_when_txo_header_is_missing_biff5() {
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
    fn falls_back_to_decoding_multiple_continue_fragments_when_txo_header_is_missing() {
        // When the TXO header is truncated/missing, we still want best-effort recovery of text
        // that spans multiple CONTINUE records.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &[]),
            continue_text_ascii("Hel"),
            continue_text_ascii("lo"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, .. } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn falls_back_to_decoding_continue_fragments_when_txo_header_is_missing_and_a_later_fragment_is_missing_flags_byte(
    ) {
        // Similar to `parses_biff8_txo_text_when_a_later_continue_fragment_is_missing_flags_byte`,
        // but the TXO header is also missing/truncated, forcing us down the fallback decoder path.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &[]),
            continue_text_compressed_bytes(b"He"),
            // Second fragment omits the BIFF8 flags byte.
            record(records::RECORD_CONTINUE, b"llo"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]),
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
        assert!(
            warnings.iter().any(|w| w.contains("missing BIFF8 flags byte")),
            "expected missing-flags warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn fallback_decode_stops_before_formatting_runs_when_txo_header_is_missing() {
        // When the TXO header is missing, the best-effort fallback decoder should still avoid
        // interpreting formatting run bytes as text.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &[]),
            continue_text_ascii("Hello"),
            // Formatting run bytes: [ich=0][ifnt=1]. These bytes do *not* have the leading
            // continued-string flags byte, so the fallback decoder must stop before decoding them.
            record(records::RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]),
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
    fn does_not_decode_biff8_formatting_runs_as_text_when_cb_runs_is_missing() {
        // If the TXO header is truncated such that `cbRuns` is missing, but `cchText` is still
        // present and larger than the actual text bytes, we should stop before decoding the
        // formatting run bytes as text.
        let mut txo_header = vec![0u8; 8];
        txo_header[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&10u16.to_le_bytes());

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_header),
            continue_text_ascii("Hello"),
            // Formatting run bytes: [ich=0][ifnt=1] (no leading flags byte).
            record(records::RECORD_CONTINUE, &[0x00, 0x00, 0x01, 0x00]),
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
    fn does_not_decode_biff8_formatting_runs_appended_to_text_fragment_when_cb_runs_is_missing() {
        // Similar to `does_not_decode_biff8_formatting_runs_as_text_when_cb_runs_is_missing`, but
        // the formatting runs are appended to the end of the *text* bytes within the same CONTINUE
        // fragment (after the 1-byte fHighByte flag).
        let mut txo_header = vec![0u8; 8];
        txo_header[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
            .copy_from_slice(&10u16.to_le_bytes());

        let mut mixed = Vec::new();
        mixed.extend_from_slice(b"Hello");
        // Formatting run bytes: [ich=0][ifnt=1].
        mixed.extend_from_slice(&[0x00, 0x00, 0x01, 0x00]);

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_header),
            continue_text_compressed_bytes(&mixed),
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
    fn does_not_decode_biff8_formatting_runs_with_large_ich_values_as_text_when_cb_runs_is_missing() {
        // Similar to `does_not_decode_biff8_formatting_runs_as_text_when_cb_runs_is_missing`, but
        // use a formatting-run payload with multiple entries where most `ich` values exceed 255
        // (high byte non-zero). The formatting-run detection heuristic should still catch this.
        const TEXT_LEN: usize = 300;
        const CCH_TEXT: u16 = 310; // larger than available text chars, so we'd otherwise decode run bytes

        let mut txo_header = vec![0u8; 8];
        txo_header[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2].copy_from_slice(&CCH_TEXT.to_le_bytes());

        let text_bytes = vec![b'A'; TEXT_LEN];

        // Formatting runs: 3 entries (12 bytes):
        // - ich=0, ifnt=1
        // - ich=256, ifnt=1
        // - ich=300, ifnt=1
        let runs: [u8; 12] = [
            0x00, 0x00, 0x01, 0x00, // 0
            0x00, 0x01, 0x01, 0x00, // 256
            0x2C, 0x01, 0x01, 0x00, // 300
        ];

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            record(RECORD_TXO, &txo_header),
            continue_text_compressed_bytes(&text_bytes),
            record(records::RECORD_CONTINUE, &runs),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "A".repeat(TEXT_LEN));
        assert!(
            warnings.iter().any(|w| w.contains("truncated text")),
            "expected truncation warning; warnings={warnings:?}"
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hi");
    }

    #[test]
    fn parses_unicode_text_split_mid_code_unit_across_continue_records() {
        // Some malformed `.xls` files split UTF-16LE code units across CONTINUE boundaries (odd byte
        // counts per fragment). Best-effort decoding should still recover the intended character.
        //
        // '€' (U+20AC) is 0xAC 0x20 in UTF-16LE. Split as 0xAC + 0x20 across two CONTINUE records.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(1),
            record(records::RECORD_CONTINUE, &[0x01, 0xAC]),
            record(records::RECORD_CONTINUE, &[0x01, 0x20]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "€");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
    fn parses_txo_text_when_cchtext_is_zero_but_continue_has_text() {
        // Some files report `cchText=0` in the TXO header even though the continuation area still
        // contains the text. Best-effort decoding should still recover it.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(0),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, .. } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn parses_txo_text_when_cchtext_is_zero_even_when_cb_runs_is_present() {
        // Some files report `cchText=0` but still set `cbRuns` (typically 4). Ensure we don't
        // accidentally treat `cbRuns` as an alternate text-length field and truncate the text.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text_and_cb_runs(0, 4),
            continue_text_ascii("Hello"),
            // Formatting runs CONTINUE payload (dummy bytes, no leading flags byte).
            record(records::RECORD_CONTINUE, &[0x00u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, .. } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].text, "Hello");
    }

    #[test]
    fn trims_trailing_nuls_when_using_fallback_txo_decode() {
        // Excel sometimes NUL-terminates TXO text. Ensure we trim trailing terminators even when
        // we need to fall back to decoding based on the continuation payload.
        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(0),
            continue_text_ascii("Hello\0\0"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, .. } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "Hi А");
    }

    #[test]
    fn parses_biff5_txo_text_with_multibyte_codepage_split_across_continue_records() {
        // In Shift-JIS (codepage 932), '\u{3042}' ('あ') is encoded as 0x82 0xA0. Ensure we decode
        // across CONTINUE boundaries without corrupting multibyte sequences.
        let sjis = [0x82u8, 0xA0u8];
        let cch_text = sjis.len() as u16;

        let stream = [
            bof_biff5(),
            note_biff5(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(cch_text),
            continue_text_biff5(&sjis[..1]),
            continue_text_biff5(&sjis[1..]),
            // Formatting CONTINUE payload (dummy bytes).
            continue_text_biff5(&[0u8; 4]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff5, 932).expect("parse");
        assert!(
            warnings.is_empty(),
            "unexpected warnings: {warnings:?}"
        );
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "\u{3042}");
    }

    #[test]
    fn parses_biff8_txo_text_with_multibyte_codepage_split_across_continue_records() {
        // Some `.xls` files appear to store BIFF8 TXO comment text as 8-bit codepage bytes even
        // when using a multibyte codepage like Shift-JIS (932). Ensure we decode across CONTINUE
        // boundaries without corrupting multibyte sequences.
        let sjis = [0x82u8, 0xA0u8]; // 'あ'
        let cch_text = sjis.len() as u16;

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(cch_text),
            continue_text_compressed_bytes(&sjis[..1]),
            continue_text_compressed_bytes(&sjis[1..]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 932).expect("parse");
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "\u{3042}");
    }

    #[test]
    fn parses_biff8_txo_text_with_surrogate_pair_split_across_continue_records() {
        // U+1F600 GRINNING FACE 😀 is encoded in UTF-16 as the surrogate pair D83D DE00. Ensure we
        // can handle the surrogate pair being split across CONTINUE fragments.
        let cch_text = 2u16; // two UTF-16 code units

        let stream = [
            bof(),
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_cch_text(cch_text),
            // High surrogate (D83D).
            record(records::RECORD_CONTINUE, &[0x01, 0x3D, 0xD8]),
            // Low surrogate (DE00).
            record(records::RECORD_CONTINUE, &[0x01, 0x00, 0xDE]),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
        assert_eq!(notes[0].text, "\u{1F600}");
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
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].author, "Alice");
    }

    #[test]
    fn caps_notes_per_sheet() {
        fn append_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
            out.extend_from_slice(&id.to_le_bytes());
            out.extend_from_slice(&(data.len() as u16).to_le_bytes());
            out.extend_from_slice(data);
        }

        let total_notes = MAX_NOTES_PER_SHEET + 10;
        // Approximate capacity: ~73 bytes per (NOTE+OBJ+TXO+CONTINUE) group.
        let mut stream = Vec::new();

        append_record(&mut stream, records::RECORD_BOF_BIFF8, &[0u8; 16]);

        for i in 0..total_notes {
            let row = i as u16;
            let obj_id = (i as u16).wrapping_add(1);

            // NOTE record payload (11 bytes for a 1-byte author).
            let mut note_payload = [0u8; 11];
            note_payload[0..2].copy_from_slice(&row.to_le_bytes());
            note_payload[2..4].copy_from_slice(&0u16.to_le_bytes()); // col
            // Write `obj_id` into both adjacent fields for robustness across parser variations.
            note_payload[4..6].copy_from_slice(&obj_id.to_le_bytes());
            note_payload[6..8].copy_from_slice(&obj_id.to_le_bytes());
            // BIFF8 ShortXLUnicodeString author: "A" (compressed).
            note_payload[8] = 1; // length
            note_payload[9] = 0; // flags (compressed)
            note_payload[10] = b'A';
            append_record(&mut stream, RECORD_NOTE, &note_payload);

            // OBJ record payload containing ftCmo with idObj.
            let mut obj_payload = [0u8; 26];
            obj_payload[0..2].copy_from_slice(&OBJ_SUBRECORD_FT_CMO.to_le_bytes());
            obj_payload[2..4].copy_from_slice(&0x0012u16.to_le_bytes()); // ftCmo size
            obj_payload[4..6].copy_from_slice(&0u16.to_le_bytes()); // ot
            obj_payload[6..8].copy_from_slice(&obj_id.to_le_bytes()); // idObj
            // Remaining bytes are zero (ftCmo tail + optional ftEnd).
            append_record(&mut stream, RECORD_OBJ, &obj_payload);

            // TXO record header with `cchText` at offset 6.
            let mut txo_payload = [0u8; 18];
            txo_payload[TXO_TEXT_LEN_OFFSET..TXO_TEXT_LEN_OFFSET + 2]
                .copy_from_slice(&1u16.to_le_bytes());
            append_record(&mut stream, RECORD_TXO, &txo_payload);

            // CONTINUE payload containing a single compressed character.
            let continue_payload = [0u8, b'x'];
            append_record(&mut stream, records::RECORD_CONTINUE, &continue_payload);
        }

        append_record(&mut stream, records::RECORD_EOF, &[]);

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes(&stream, 0, BiffVersion::Biff8, 1252).expect("parse");
        assert_eq!(notes.len(), MAX_NOTES_PER_SHEET);
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("capped") && w.contains("NOTE")),
            "expected truncation warning; warnings={warnings:?}"
        );
    }

    #[test]
    fn sheet_notes_scan_stops_after_record_cap() {
        let record_cap = 10usize;

        let stream = [
            bof(),
            // Exceed the record-scan cap with junk records.
            (0..(record_cap + 10))
                .flat_map(|_| record(0x1234, &[]))
                .collect::<Vec<u8>>(),
            // This note would be parsed if we scanned further.
            note(0, 0, 1, "Alice"),
            obj_with_id(1),
            txo_with_text("Hello"),
            continue_text_ascii("Hello"),
            eof(),
        ]
        .concat();

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes_with_record_cap(&stream, 0, BiffVersion::Biff8, 1252, record_cap)
                .expect("parse");
        assert!(notes.is_empty(), "expected no notes, got {notes:?}");
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("worksheet notes")),
            "expected record-cap warning, got {warnings:?}"
        );
    }

    #[test]
    fn sheet_notes_record_cap_warning_is_emitted_even_when_warning_buffer_is_full() {
        let record_cap = MAX_WARNINGS_PER_SHEET + 10;

        let mut stream = Vec::new();
        stream.extend_from_slice(&bof());

        // Fill the warning buffer with TXO records missing preceding OBJ ids.
        for _ in 0..(MAX_WARNINGS_PER_SHEET + 10) {
            stream.extend_from_slice(&record(RECORD_TXO, &[]));
        }

        // Exceed the record-scan cap.
        for _ in 0..(record_cap + 10) {
            stream.extend_from_slice(&record(0x1234, &[]));
        }
        stream.extend_from_slice(&eof());

        let ParsedSheetNotes { notes, warnings } =
            parse_biff_sheet_notes_with_record_cap(&stream, 0, BiffVersion::Biff8, 1252, record_cap)
                .expect("parse");
        assert!(notes.is_empty());
        assert_eq!(
            warnings.len(),
            MAX_WARNINGS_PER_SHEET,
            "warnings should remain capped; warnings={warnings:?}"
        );
        assert!(
            warnings
                .iter()
                .any(|w| w.contains("too many BIFF records") && w.contains("worksheet notes")),
            "expected forced record-cap warning, got {warnings:?}"
        );
    }
}
