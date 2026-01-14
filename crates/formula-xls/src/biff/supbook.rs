//! BIFF8 `SUPBOOK` (0x01AE) / `EXTERNNAME` (0x0023) parsing.
//!
//! External workbook references in BIFF8 3D reference tokens (`PtgRef3d`, `PtgArea3d`) are
//! resolved via the workbook-global `EXTERNSHEET` table (XTI structures). Each XTI entry points at
//! a `SUPBOOK` record (`iSupBook`) that contains the referenced workbook name and sheet names.
//!
//! This module provides a minimal, best-effort parser for:
//! - `SUPBOOK` records (including continued BIFF8 strings across `CONTINUE` records)
//! - `EXTERNNAME` records (captured and associated with the preceding `SUPBOOK`)
//!
//! The `.xls` importer must never hard-fail due to malformed external reference metadata. All
//! parsing is best-effort and any issues are reported as warnings.

#![allow(dead_code)]

use super::{records, strings};

/// BIFF8 `SUPBOOK` record id.
///
/// See [MS-XLS] 2.4.271 (SUPBOOK).
const RECORD_SUPBOOK: u16 = 0x01AE;

/// BIFF8 `EXTERNNAME` record id.
///
/// See [MS-XLS] 2.4.106 (EXTERNNAME).
const RECORD_EXTERNNAME: u16 = 0x0023;

/// Maximum number of warnings retained while scanning for `SUPBOOK` / `EXTERNNAME` metadata.
///
/// Corrupt workbook streams can contain huge numbers of malformed records; capping warning growth
/// prevents unbounded memory usage while still surfacing that something went wrong.
const MAX_SUPBOOK_WARNINGS: usize = 200;

const SUPBOOK_WARNINGS_SUPPRESSED_MSG: &str =
    "too many SUPBOOK/EXTERNNAME warnings; further warnings suppressed";

// BIFF8 string option flags used by XLUnicodeString.
// See [MS-XLS] 2.5.268.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SupBookKind {
    /// Internal workbook marker (`virtPath` is a single control character).
    Internal,
    /// An external workbook reference with a file name/path and sheet name list.
    ExternalWorkbook,
    /// An add-in function library (XLL) or other special supbook type.
    Other,
}

#[derive(Debug, Clone)]
pub(crate) struct SupBookInfo {
    pub(crate) ctab: u16,
    /// Raw `virtPath` string.
    pub(crate) virt_path: String,
    pub(crate) kind: SupBookKind,
    /// Best-effort extracted workbook base name (without path), used for formula rendering.
    ///
    /// Only populated for `ExternalWorkbook` supbooks.
    pub(crate) workbook_name: Option<String>,
    /// Sheet names stored in `SUPBOOK` (for external workbooks).
    pub(crate) sheet_names: Vec<String>,
    /// External names (`EXTERNNAME`) belonging to this supbook, in record order.
    pub(crate) extern_names: Vec<String>,
}

impl SupBookInfo {
    pub(crate) fn is_internal(&self) -> bool {
        self.kind == SupBookKind::Internal
    }
}

#[derive(Debug, Default, Clone)]
pub(crate) struct SupBookTable {
    pub(crate) supbooks: Vec<SupBookInfo>,
    pub(crate) warnings: Vec<String>,
}

fn push_warning(out: &mut SupBookTable, msg: String) {
    if out.warnings.len() < MAX_SUPBOOK_WARNINGS {
        out.warnings.push(msg);
        return;
    }

    // Cap reached: keep the list bounded and emit a single suppression marker.
    if out.warnings.len() > MAX_SUPBOOK_WARNINGS {
        out.warnings.truncate(MAX_SUPBOOK_WARNINGS);
    }
    if out
        .warnings
        .last()
        .is_some_and(|w| w == SUPBOOK_WARNINGS_SUPPRESSED_MSG)
    {
        return;
    }

    // Keep `warnings.len() == MAX_SUPBOOK_WARNINGS` by overwriting the last entry.
    if let Some(last) = out.warnings.last_mut() {
        *last = SUPBOOK_WARNINGS_SUPPRESSED_MSG.to_string();
    }
}

pub(crate) fn parse_biff8_supbook_table(workbook_stream: &[u8], codepage: u16) -> SupBookTable {
    let mut out = SupBookTable::default();

    let allows_continuation = |id: u16| id == RECORD_SUPBOOK || id == RECORD_EXTERNNAME;
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    let mut current_supbook: Option<usize> = None;

    for record in iter {
        let record = match record {
            Ok(record) => record,
            Err(err) => {
                push_warning(
                    &mut out,
                    format!("malformed BIFF record while scanning for SUPBOOK/EXTERNNAME: {err}"),
                );
                break;
            }
        };

        // Stop scanning at the start of the next substream (worksheet BOF), even if the workbook
        // globals are missing the expected EOF record.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_SUPBOOK => {
                let (info, warnings) = parse_supbook_record(&record, codepage);
                for warning in warnings {
                    push_warning(&mut out, warning);
                }
                out.supbooks.push(info);
                current_supbook = Some(out.supbooks.len().saturating_sub(1));
            }
            RECORD_EXTERNNAME => {
                let Some(idx) = current_supbook else {
                    push_warning(
                        &mut out,
                        format!(
                            "EXTERNNAME record at offset {} without preceding SUPBOOK",
                            record.offset
                        ),
                    );
                    continue;
                };

                match parse_externname_record(&record, codepage) {
                    Ok(name) => out.supbooks[idx].extern_names.push(name),
                    Err(err) => push_warning(
                        &mut out,
                        format!(
                            "failed to parse EXTERNNAME record at offset {}: {err}",
                            record.offset
                        ),
                    ),
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    out
}

fn parse_supbook_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
) -> (SupBookInfo, Vec<String>) {
    let mut warnings = Vec::new();

    // Best-effort handling for minimal internal SUPBOOK marker payloads.
    //
    // Some producers emit an "internal references" SUPBOOK record as a 4-byte payload:
    //   [ctab: u16][marker: u16]
    // where `marker == 0x0401` (little-endian bytes 0x01 0x04).
    //
    // This is not an XLUnicodeString `virtPath` and will not parse via the standard string path.
    // Treat it as an internal workbook reference.
    let raw = record.data.as_ref();
    if raw.len() == 4 {
        let ctab = u16::from_le_bytes([raw[0], raw[1]]);
        let marker = u16::from_le_bytes([raw[2], raw[3]]);
        if marker == 0x0401 {
            return (
                SupBookInfo {
                    ctab,
                    virt_path: "\u{0001}\u{0004}".to_string(),
                    kind: SupBookKind::Internal,
                    workbook_name: None,
                    sheet_names: Vec::new(),
                    extern_names: Vec::new(),
                },
                warnings,
            );
        }
    }

    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    let ctab = match cursor.read_u16_le() {
        Ok(v) => v,
        Err(err) => {
            warnings.push(format!(
                "truncated SUPBOOK record at offset {}: {err}",
                record.offset
            ));
            return (
                SupBookInfo {
                    ctab: 0,
                    virt_path: String::new(),
                    kind: SupBookKind::Other,
                    workbook_name: None,
                    sheet_names: Vec::new(),
                    extern_names: Vec::new(),
                },
                warnings,
            );
        }
    };

    let virt_path = match cursor.read_biff8_unicode_string(codepage) {
        Ok(v) => v,
        Err(err) => {
            warnings.push(format!(
                "failed to decode SUPBOOK virtPath at offset {}: {err}",
                record.offset
            ));
            String::new()
        }
    };

    let kind = if is_internal_virt_path(&virt_path) {
        SupBookKind::Internal
    } else if is_addin_virt_path(&virt_path) {
        SupBookKind::Other
    } else {
        SupBookKind::ExternalWorkbook
    };

    let workbook_name = (kind == SupBookKind::ExternalWorkbook)
        .then(|| workbook_name_from_virt_path(&virt_path))
        .filter(|s| !s.is_empty());

    // External workbook supbooks store ctab sheet names after virtPath.
    let mut sheet_names: Vec<String> = Vec::new();
    if kind == SupBookKind::ExternalWorkbook {
        // Defend against absurd `ctab` values from corrupt files.
        const MAX_SHEETS: u16 = 4096;
        let sheet_count = if ctab > MAX_SHEETS {
            warnings.push(format!(
                "SUPBOOK record at offset {} has implausible ctab={ctab}; capping to {MAX_SHEETS}",
                record.offset
            ));
            MAX_SHEETS
        } else {
            ctab
        };

        sheet_names.reserve(sheet_count as usize);

        for sheet_idx in 0..sheet_count {
            match cursor.read_biff8_unicode_string(codepage) {
                Ok(name) => sheet_names.push(name),
                Err(err) => {
                    warnings.push(format!(
                        "failed to decode SUPBOOK sheet name {sheet_idx} at offset {}: {err}",
                        record.offset
                    ));
                    break;
                }
            }
        }
    }

    (
        SupBookInfo {
            ctab,
            virt_path,
            kind,
            workbook_name,
            sheet_names,
            extern_names: Vec::new(),
        },
        warnings,
    )
}

fn is_internal_virt_path(virt_path: &str) -> bool {
    // There are multiple conventions in the wild for internal marker strings. Excel typically uses
    // a single 0x0001 character, but some writers appear to use NUL or a multi-character marker.
    //
    // Be permissive about trailing NUL padding: some producers write marker strings with one or
    // more trailing `\0` characters.
    let trimmed = virt_path.trim_end_matches('\0');
    // If the string is entirely NULs (e.g. `"\0"`), it represents the internal marker, but treat
    // an actually-empty string as "unknown" (it may come from a decode failure).
    if trimmed.is_empty() {
        return !virt_path.is_empty();
    }
    trimmed == "\u{0001}" || trimmed == "\u{0000}" || trimmed == "\u{0001}\u{0004}"
}

fn is_addin_virt_path(virt_path: &str) -> bool {
    // Excel uses a single 0x0002 character for add-in references.
    virt_path.trim_end_matches('\0') == "\u{0002}"
}

fn workbook_name_from_virt_path(virt_path: &str) -> String {
    // Best-effort conversion:
    // - strip embedded NULs
    // - take basename after path separators
    // - strip Excel-style wrapper brackets if present (but preserve literal `[` / `]` characters
    //   in actual workbook names)
    let without_nuls = virt_path.replace('\0', "");

    let trimmed_full = without_nuls.trim();
    let has_full_wrapper = trimmed_full.starts_with('[') && trimmed_full.ends_with(']');

    let basename = trimmed_full
        .rsplit(['\\', '/'])
        .next()
        .unwrap_or(trimmed_full);

    let trimmed = basename.trim();
    let has_basename_wrapper = trimmed.starts_with('[') && trimmed.ends_with(']');

    // Be permissive about bracket placement: some producers wrap the full path in brackets
    // (`[C:\\path\\Book.xlsx]`), which means we might lose the opening `[` when taking the basename.
    //
    // Only strip wrapper brackets when the input appears to be wrapper-bracketed, so we don't
    // drop legitimate leading `[` / trailing `]` characters in workbook names.
    let mut inner = trimmed;
    if has_full_wrapper || has_basename_wrapper {
        inner = inner.strip_prefix('[').unwrap_or(inner);
        inner = inner.strip_suffix(']').unwrap_or(inner);
    }

    inner.to_string()
}

fn parse_externname_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
) -> Result<String, String> {
    // Best-effort EXTERNNAME parsing.
    //
    // The record structure is complex and varies depending on flags (add-in, OLE/DDE, etc). For
    // the purposes of `PtgNameX` rendering we only need the name text. Most common producers store
    // an `XLUnicodeStringNoCch` after a small fixed header:
    //   [grbit: u16][reserved: u32][cch: u8][rgchName: XLUnicodeStringNoCch]
    //
    // Since we cannot rely on all variants, we implement a conservative heuristic:
    // - attempt the common header layout first
    // - if that fails, attempt to locate a plausible `XLUnicodeString` at the end of the record
    //
    // If decoding fails, callers should treat the name as unavailable and fall back to `#REF!`.
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // Try common layout: [grbit: u16][reserved: u32][cch: u8]...
    let _grbit = cursor.read_u16_le()?;
    let _reserved = cursor.read_u32_le()?;
    let cch = cursor.read_u8()? as usize;
    let name = cursor.read_biff8_unicode_string_no_cch(cch, codepage)?;
    if !name.is_empty() {
        return Ok(name);
    }

    // Fallback: scan the first fragment for a trailing XLUnicodeString (u16 cch + flags).
    // This is intentionally best-effort and may return an empty string.
    let raw = record.data.as_ref();
    if raw.len() < 3 {
        return Err("truncated EXTERNNAME record".to_string());
    }
    if let Some(name) = strings::parse_biff8_unicode_string_best_effort(raw, codepage) {
        return Ok(name);
    }

    Err("failed to decode EXTERNNAME string".to_string())
}

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
            .map(|f| f.len().saturating_sub(self.offset))
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

    fn read_u32_le(&mut self) -> Result<u32, String> {
        let b0 = self.read_u8()?;
        let b1 = self.read_u8()?;
        let b2 = self.read_u8()?;
        let b3 = self.read_u8()?;
        Ok(u32::from_le_bytes([b0, b1, b2, b3]))
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

    fn skip_bytes(&mut self, mut n: usize) -> Result<(), String> {
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            self.offset += take;
            n -= take;
        }
        Ok(())
    }

    fn advance_fragment_in_biff8_string(&mut self, is_unicode: &mut bool) -> Result<(), String> {
        self.advance_fragment()?;
        // When a BIFF8 string spans a CONTINUE boundary, Excel inserts a 1-byte option flags prefix
        // at the start of the continued fragment. The only relevant bit is `fHighByte` (unicode vs
        // compressed).
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
        let mut out = Vec::with_capacity(n);
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
                // Continuing character bytes into a new CONTINUE fragment: first byte is option
                // flags for the continued segment (fHighByte).
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
                let mut u16s = Vec::with_capacity(take_chars);
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

    fn read_biff8_unicode_string_no_cch(
        &mut self,
        cch: usize,
        codepage: u16,
    ) -> Result<String, String> {
        // XLUnicodeStringNoCch [MS-XLS 2.5.277] (used by NAME/EXTERNNAME).
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
                let mut u16s = Vec::with_capacity(take_chars);
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

    fn xl_unicode_string_compressed(s: &str) -> Vec<u8> {
        let bytes = s.as_bytes();
        let cch: u16 = bytes.len().try_into().expect("len fits u16");
        [cch.to_le_bytes().to_vec(), vec![0u8], bytes.to_vec()].concat()
    }

    fn externname_record_payload_compressed(name: &str) -> Vec<u8> {
        // Best-effort EXTERNNAME payload matching the common layout parsed by `parse_externname_record`:
        //   [grbit: u16][reserved: u32][cch: u8][XLUnicodeStringNoCch]
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        payload.extend_from_slice(&0u32.to_le_bytes()); // reserved
        payload.push(name.len() as u8); // cch
        payload.push(0); // flags (compressed)
        payload.extend_from_slice(name.as_bytes());
        payload
    }

    #[test]
    fn parses_external_supbook_with_sheet_list() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes()); // ctab
        payload.extend_from_slice(&xl_unicode_string_compressed("Book2.xlsx"));
        payload.extend_from_slice(&xl_unicode_string_compressed("Sheet1"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::ExternalWorkbook);
        assert_eq!(sb.virt_path, "Book2.xlsx");
        assert_eq!(sb.workbook_name.as_deref(), Some("Book2.xlsx"));
        assert_eq!(sb.sheet_names, vec!["Sheet1".to_string()]);
        assert!(sb.extern_names.is_empty());
    }

    #[test]
    fn parses_continued_supbook_workbook_name() {
        let full_name = "ABCDEFGHIJ";

        let mut full_payload = Vec::new();
        full_payload.extend_from_slice(&0u16.to_le_bytes()); // ctab=0 (no sheet names)
        full_payload.extend_from_slice(&xl_unicode_string_compressed(full_name));

        // Split in the middle of the character bytes ("ABCD" | "EFGHIJ").
        let split_at = 2 + 2 + 1 + 4;
        let first = &full_payload[..split_at];
        let rest = &full_payload[split_at..];

        // BIFF8 inserts a single option-flags byte at the start of a continued string segment.
        let second = [vec![0u8], rest.to_vec()].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, first),
            record(records::RECORD_CONTINUE, &second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::ExternalWorkbook);
        assert_eq!(sb.virt_path, full_name);
        assert_eq!(sb.workbook_name.as_deref(), Some(full_name));
        assert!(sb.sheet_names.is_empty());
    }

    #[test]
    fn parses_internal_supbook_marker_without_sheet_names() {
        // Internal SUPBOOK: ctab=sheet count, virtPath is the single 0x01 marker character.
        let mut payload = Vec::new();
        payload.extend_from_slice(&3u16.to_le_bytes()); // ctab (sheet count)
        payload.extend_from_slice(&xl_unicode_string_compressed("\u{0001}"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::Internal);
        assert_eq!(sb.virt_path, "\u{0001}");
        assert_eq!(sb.workbook_name, None);
        assert!(sb.sheet_names.is_empty());
        assert!(sb.extern_names.is_empty());
    }

    #[test]
    fn parses_internal_supbook_marker_0401_without_sheet_names() {
        // Internal SUPBOOK via minimal 4-byte payload:
        //   [ctab: u16][marker: u16=0x0401]
        let mut payload = Vec::new();
        payload.extend_from_slice(&3u16.to_le_bytes()); // ctab (sheet count)
        payload.extend_from_slice(&0x0401u16.to_le_bytes()); // marker

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::Internal);
        assert_eq!(sb.virt_path, "\u{0001}\u{0004}");
        assert_eq!(sb.workbook_name, None);
        assert!(sb.sheet_names.is_empty());
        assert!(sb.extern_names.is_empty());
    }

    #[test]
    fn parses_internal_supbook_marker_with_trailing_nuls() {
        // Some producers emit internal marker strings with trailing NUL padding.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        payload.extend_from_slice(&xl_unicode_string_compressed("\u{0001}\u{0000}"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::Internal);
    }

    #[test]
    fn parses_addin_supbook_marker_with_trailing_nuls() {
        // Excel uses a single 0x0002 marker for add-in references, but some producers add NULs.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        payload.extend_from_slice(&xl_unicode_string_compressed("\u{0002}\u{0000}"));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::Other);
    }

    #[test]
    fn extracts_workbook_name_from_path_and_brackets() {
        let virt_path = "C:\\tmp\\[Book2.xlsx]\u{0000}";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        payload.extend_from_slice(&xl_unicode_string_compressed(virt_path));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::ExternalWorkbook);
        assert_eq!(sb.workbook_name.as_deref(), Some("Book2.xlsx"));
    }

    #[test]
    fn extracts_workbook_name_from_bracketed_path() {
        // Some producers wrap the full path in brackets (`[C:\tmp\Book2.xlsx]`). When we take the
        // basename after the last path separator, we lose the leading `[` but keep the trailing
        // `]` (e.g. `Book2.xlsx]`). We should still normalize to `Book2.xlsx`.
        let virt_path = "[C:\\tmp\\Book2.xlsx]\u{0000}";

        let mut payload = Vec::new();
        payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        payload.extend_from_slice(&xl_unicode_string_compressed(virt_path));

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        let sb = &parsed.supbooks[0];
        assert_eq!(sb.kind, SupBookKind::ExternalWorkbook);
        assert_eq!(sb.workbook_name.as_deref(), Some("Book2.xlsx"));
    }

    #[test]
    fn preserves_literal_brackets_in_workbook_names() {
        // Workbook names may contain literal `[` / `]` characters. Those are distinct from the
        // Excel-style wrapper brackets used in some SUPBOOK virtPath strings.
        //
        // When the virtPath is a plain workbook name (no wrapper pair), preserve literal bracket
        // characters rather than stripping them.
        for (virt_path, expected) in [
            ("[LeadingBracket.xlsx", "[LeadingBracket.xlsx"),
            ("Book2.xlsx]", "Book2.xlsx]"),
        ] {
            let mut payload = Vec::new();
            payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
            payload.extend_from_slice(&xl_unicode_string_compressed(virt_path));

            let stream = [
                record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
                record(RECORD_SUPBOOK, &payload),
                record(records::RECORD_EOF, &[]),
            ]
            .concat();

            let parsed = parse_biff8_supbook_table(&stream, 1252);
            assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
            assert_eq!(parsed.supbooks.len(), 1);
            let sb = &parsed.supbooks[0];
            assert_eq!(sb.kind, SupBookKind::ExternalWorkbook);
            assert_eq!(
                sb.workbook_name.as_deref(),
                Some(expected),
                "virt_path={virt_path:?}"
            );
        }
    }

    #[test]
    fn parses_externname_records_after_supbook() {
        let mut sb_payload = Vec::new();
        sb_payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        sb_payload.extend_from_slice(&xl_unicode_string_compressed("Book2.xlsx"));

        let externname_payload = externname_record_payload_compressed("MyName");

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &sb_payload),
            record(RECORD_EXTERNNAME, &externname_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        assert_eq!(parsed.supbooks[0].extern_names, vec!["MyName".to_string()]);
    }

    #[test]
    fn parses_continued_externname_strings() {
        let mut sb_payload = Vec::new();
        sb_payload.extend_from_slice(&0u16.to_le_bytes()); // ctab
        sb_payload.extend_from_slice(&xl_unicode_string_compressed("Book2.xlsx"));

        let full = externname_record_payload_compressed("ABCDEFG");
        // Split after "ABC" so the remaining chars ("DEFG") are in a CONTINUE fragment.
        let split_at = 2 + 4 + 1 + 1 + 3;
        let first = &full[..split_at];
        let rest = &full[split_at..];

        // BIFF8 inserts an option flags byte at the start of each continued string fragment.
        let second = [vec![0u8], rest.to_vec()].concat();

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_SUPBOOK, &sb_payload),
            record(RECORD_EXTERNNAME, first),
            record(records::RECORD_CONTINUE, &second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert!(parsed.warnings.is_empty(), "warnings={:?}", parsed.warnings);
        assert_eq!(parsed.supbooks.len(), 1);
        assert_eq!(parsed.supbooks[0].extern_names, vec!["ABCDEFG".to_string()]);
    }

    #[test]
    fn caps_warning_growth_for_externname_records_without_supbook() {
        // Corrupt streams can contain a long run of EXTERNNAME records before any SUPBOOK.
        // We should cap warning growth to avoid unbounded allocations.
        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            (0..(MAX_SUPBOOK_WARNINGS + 50))
                .flat_map(|_| record(RECORD_EXTERNNAME, &[]))
                .collect::<Vec<u8>>(),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let parsed = parse_biff8_supbook_table(&stream, 1252);
        assert_eq!(parsed.warnings.len(), MAX_SUPBOOK_WARNINGS);
        assert!(
            parsed
                .warnings
                .iter()
                .any(|w| w == SUPBOOK_WARNINGS_SUPPRESSED_MSG),
            "warnings={:?}",
            parsed.warnings
        );
    }

    #[test]
    fn parses_biff8_unicode_string_continued_richtext_with_crun_split_across_fragments() {
        let s = "ABCDE";
        let rg_run = [0x11u8, 0x22, 0x33, 0x44];

        // First fragment contains the header and the first byte of cRun.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&(s.len() as u16).to_le_bytes());
        frag1.push(STR_FLAG_RICH_TEXT); // flags (compressed + rich text)
        frag1.push(0x01); // cRun low byte (cRun=1)

        // Continuation fragment begins with option flags byte (fHighByte), then remaining bytes:
        // cRun high byte + character bytes + rgRun bytes.
        let mut frag2 = Vec::new();
        frag2.push(0); // continued segment compressed
        frag2.push(0x00); // cRun high byte
        frag2.extend_from_slice(s.as_bytes());
        frag2.extend_from_slice(&rg_run);

        let fragments: [&[u8]; 2] = [&frag1, &frag2];
        let mut cursor = FragmentCursor::new(&fragments, 0, 0);
        let out = cursor.read_biff8_unicode_string(1252).expect("parse");
        assert_eq!(out, s);
    }

    #[test]
    fn parses_biff8_unicode_string_continued_ext_payload_split_preserves_following_string() {
        // Two back-to-back XLUnicodeString values in a single fragment stream.
        // The first has `fExtSt=1` and its ext payload is split across fragments; the second should
        // still be parsed correctly.
        let s1 = "abc";
        let ext = [0xDEu8, 0xAD, 0xBE, 0xEF];
        let s2 = "Z";

        // First fragment: header + char bytes + first 2 ext bytes.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&(s1.len() as u16).to_le_bytes());
        frag1.push(STR_FLAG_EXT); // flags (compressed + ext)
        frag1.extend_from_slice(&(ext.len() as u32).to_le_bytes()); // cbExtRst
        frag1.extend_from_slice(s1.as_bytes());
        frag1.extend_from_slice(&ext[..2]);

        // Continuation fragment: cont_flags + remaining ext bytes + second string.
        let mut frag2 = Vec::new();
        frag2.push(0); // continued segment compressed
        frag2.extend_from_slice(&ext[2..]);
        frag2.extend_from_slice(&(s2.len() as u16).to_le_bytes());
        frag2.push(0); // flags (compressed)
        frag2.extend_from_slice(s2.as_bytes());

        let fragments: [&[u8]; 2] = [&frag1, &frag2];
        let mut cursor = FragmentCursor::new(&fragments, 0, 0);
        let out1 = cursor
            .read_biff8_unicode_string(1252)
            .expect("parse first string");
        let out2 = cursor
            .read_biff8_unicode_string(1252)
            .expect("parse second string");
        assert_eq!(out1, s1);
        assert_eq!(out2, s2);
    }
}
