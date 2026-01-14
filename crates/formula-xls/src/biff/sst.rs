//! BIFF8 SST (Shared String Table) helpers.
//!
//! This module implements best-effort extraction of phonetic guide (furigana) text from BIFF8
//! shared strings.
//!
//! ## Background
//!
//! BIFF8 stores most strings in a workbook-global Shared String Table (`SST`, record `0x00FC`).
//! Worksheet string cells (`LABELSST`, record `0x00FD`) reference strings by `isst` index.
//!
//! Each SST entry is an `XLUnicodeRichExtendedString` ([MS-XLS] 2.5.296) which may carry an
//! extended-string payload (`ExtRst`) when the `fExtSt` flag (`STR_FLAG_EXT`) is set. Excel uses
//! that payload to store phonetic guide information for East Asian text (see [MS-XLS] 2.5.86
//! `ExtRst` and 2.5.196 `PhoneticInfo`).
//!
//! The `ExtRst` layout is underspecified and varies between producers. For robustness, we:
//! - parse the SST record using [`records::LogicalBiffRecordIter`] so `CONTINUE` fragments are
//!   coalesced, while still tracking physical fragment boundaries for continued strings.
//! - treat `ExtRst` parsing as best-effort. If an `ExtRst` block is malformed or uses an unknown
//!   layout, we return `None` for that SST entry instead of failing the entire import.
//!
//! Today we extract only the phonetic *text* (a Unicode string), not the per-character phonetic
//! run mapping.

use super::records;
use super::strings;

// Record ids.
const RECORD_SST: u16 = 0x00FC;

// BIFF8 string option flags used by `XLUnicodeRichExtendedString`.
// See [MS-XLS] 2.5.293 and 2.5.268.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

// ExtRst "rt" type for phonetic blocks.
//
// [MS-XLS] names this record-type field `rt` and uses `0x0001` for phonetic information.
const EXT_RST_TYPE_PHONETIC: u16 = 0x0001;

/// Parse workbook-global SST shared string entries and extract per-entry phonetic guide text.
///
/// Returns a vector mapping `sst_index -> phonetic_text`.
///
/// This is a best-effort parser:
/// - malformed SST records yield an `Err`
/// - malformed/unknown `ExtRst` layouts yield `None` for the affected SST entries
pub(crate) fn parse_biff8_sst_phonetics(
    workbook_stream: &[u8],
    codepage: u16,
) -> Result<Vec<Option<String>>, String> {
    let allows_continuation = |id: u16| id == RECORD_SST;
    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    for record in iter {
        let record = record?;

        // Stop at the next substream BOF; workbook globals start at offset 0.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_SST => return Ok(parse_sst_record_phonetics(&record, codepage)),
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(Vec::new())
}

fn parse_sst_record_phonetics(record: &records::LogicalBiffRecord<'_>, codepage: u16) -> Vec<Option<String>> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments);

    // SST record header [MS-XLS] 2.4.261:
    //   [cstTotal: u32] [cstUnique: u32] [rgb: XLUnicodeRichExtendedString[]]
    let _cst_total = match cursor.read_u32_le() {
        Ok(v) => v,
        Err(_) => return Vec::new(),
    };
    let cst_unique = match cursor.read_u32_le() {
        Ok(v) => v as usize,
        Err(_) => return Vec::new(),
    };

    let mut out: Vec<Option<String>> = Vec::with_capacity(cst_unique.min(1024));
    for _ in 0..cst_unique {
        match cursor.read_xl_unicode_rich_extended_string_phonetic(codepage) {
            Ok(v) => out.push(v),
            Err(_) => break,
        }
    }

    out
}

fn extract_phonetic_from_ext_rst(ext: &[u8], codepage: u16) -> Option<String> {
    // Prefer the [MS-XLS] `ExtRst` TLV layout:
    //
    //   ExtRst = ( [rt: u16] [cb: u16] [rgb: cb bytes] )*
    //
    // where `rt == 0x0001` indicates phonetic information (PhoneticInfo).
    let mut pos = 0usize;
    while pos + 4 <= ext.len() {
        let rt = u16::from_le_bytes([ext[pos], ext[pos + 1]]);
        let cb = u16::from_le_bytes([ext[pos + 2], ext[pos + 3]]) as usize;
        pos += 4;
        let end = match pos.checked_add(cb) {
            Some(v) => v,
            None => break,
        };
        let Some(payload) = ext.get(pos..end) else {
            break;
        };
        pos = end;

        if rt == EXT_RST_TYPE_PHONETIC {
            if let Some(s) = scan_for_embedded_unicode_string(payload, codepage) {
                return Some(s);
            }
        }
    }

    // Fallback: scan the raw ExtRst bytes for an embedded XLUnicodeString.
    scan_for_embedded_unicode_string(ext, codepage)
}

fn scan_for_embedded_unicode_string(bytes: &[u8], codepage: u16) -> Option<String> {
    let max_start = bytes.len().saturating_sub(3).min(32);
    let mut best: Option<(usize, String)> = None;

    for start in 0..=max_start {
        let Ok((mut s, _consumed)) = strings::parse_biff8_unicode_string(&bytes[start..], codepage)
        else {
            continue;
        };

        // Excel sometimes embeds NULs in BIFF strings; strip for stability.
        if s.contains('\0') {
            s.retain(|c| c != '\0');
        }
        if s.is_empty() {
            continue;
        }

        // Score by number of non-control codepoints.
        let score = s.chars().filter(|c| !c.is_control()).count();
        match best.as_ref() {
            Some((best_score, _)) if *best_score >= score => {}
            _ => best = Some((score, s)),
        }
    }

    best.map(|(_, s)| s)
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

    fn read_bytes(&mut self, n: usize) -> Result<Vec<u8>, String> {
        let mut out = Vec::with_capacity(n);
        let mut remaining = n;
        while remaining > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = remaining.min(available);
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            let end = self.offset + take;
            out.extend_from_slice(&frag[self.offset..end]);
            self.offset = end;
            remaining -= take;
        }
        Ok(out)
    }

    fn skip_biff8_char_data(&mut self, cch: usize, initial_is_unicode: bool) -> Result<(), String> {
        let mut is_unicode = initial_is_unicode;
        let mut remaining_chars = cch;

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first
                // byte is option flags for the continued segment (fHighByte).
                self.advance_fragment()?;
                let cont_flags = self.read_u8()?;
                is_unicode = (cont_flags & STR_FLAG_HIGH_BYTE) != 0;
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
            self.skip_bytes(take_bytes)?;
            remaining_chars -= take_chars;
        }

        Ok(())
    }

    fn read_xl_unicode_rich_extended_string_phonetic(
        &mut self,
        codepage: u16,
    ) -> Result<Option<String>, String> {
        // XLUnicodeRichExtendedString [MS-XLS] 2.5.296.
        let cch = self.read_u16_le()? as usize;
        let flags = self.read_u8()?;

        let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
            self.read_u16_le()? as usize
        } else {
            0
        };

        let ext_size = if flags & STR_FLAG_EXT != 0 {
            self.read_u32_le()? as usize
        } else {
            0
        };

        let is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
        self.skip_biff8_char_data(cch, is_unicode)?;

        let richtext_bytes = richtext_runs
            .checked_mul(4)
            .ok_or_else(|| "rich text run count overflow".to_string())?;
        self.skip_bytes(richtext_bytes)?;

        if ext_size == 0 {
            return Ok(None);
        }

        // Avoid pathological allocation on corrupt files.
        const MAX_EXT_RST_BYTES: usize = 1024 * 1024; // 1 MiB
        if ext_size > MAX_EXT_RST_BYTES {
            self.skip_bytes(ext_size)?;
            return Ok(None);
        }

        let ext_bytes = self.read_bytes(ext_size)?;
        Ok(extract_phonetic_from_ext_rst(&ext_bytes, codepage))
    }
}

