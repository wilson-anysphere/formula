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

    let mut out: Vec<Option<String>> = Vec::new();
    let _ = out.try_reserve_exact(cst_unique.min(1024));
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
    while let Some(header) = ext.get(pos..).and_then(|rest| rest.get(..4)) {
        let rt = u16::from_le_bytes([header[0], header[1]]);
        let cb = u16::from_le_bytes([header[2], header[3]]) as usize;
        pos = pos.checked_add(4)?;
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
        let total = n;
        let mut out = Vec::new();
        out.try_reserve_exact(total)
            .map_err(|_| "allocation failed (sst string bytes)".to_string())?;
        while n > 0 {
            if self.remaining_in_fragment() == 0 {
                self.advance_fragment_in_biff8_string(is_unicode)?;
                continue;
            }
            let available = self.remaining_in_fragment();
            let take = n.min(available);
            let frag = self
                .fragments
                .get(self.frag_idx)
                .ok_or_else(|| "unexpected end of record".to_string())?;
            let end = self.offset + take;
            out.extend_from_slice(&frag[self.offset..end]);
            self.offset = end;
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

    fn skip_biff8_char_data(&mut self, cch: usize, initial_is_unicode: bool) -> Result<(), String> {
        let mut is_unicode = initial_is_unicode;
        let mut remaining_chars = cch;

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first
                // byte is option flags for the continued segment (fHighByte).
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

        self.skip_biff8_char_data(cch, is_unicode)?;

        let richtext_bytes = richtext_runs
            .checked_mul(4)
            .ok_or_else(|| "rich text run count overflow".to_string())?;
        self.skip_biff8_string_bytes(richtext_bytes, &mut is_unicode)?;

        if ext_size == 0 {
            return Ok(None);
        }

        // Avoid pathological allocation on corrupt files.
        const MAX_EXT_RST_BYTES: usize = 1024 * 1024; // 1 MiB
        if ext_size > MAX_EXT_RST_BYTES {
            self.skip_biff8_string_bytes(ext_size, &mut is_unicode)?;
            return Ok(None);
        }

        let ext_bytes = self.read_biff8_string_bytes(ext_size, &mut is_unicode)?;
        Ok(extract_phonetic_from_ext_rst(&ext_bytes, codepage))
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

    fn xl_unicode_string_compressed(s: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(s.len() as u16).to_le_bytes());
        out.push(0); // flags (compressed)
        out.extend_from_slice(s.as_bytes());
        out
    }

    #[test]
    fn parses_sst_phonetic_with_richtext_crun_split_across_continue() {
        let main = "ABCDE";
        let phonetic = "kana";
        let ext_payload = xl_unicode_string_compressed(phonetic);
        let rg_run = [0x11u8, 0x22, 0x33, 0x44];

        // SST header + string header up through the first byte of cRun.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&1u32.to_le_bytes()); // cstTotal
        frag1.extend_from_slice(&1u32.to_le_bytes()); // cstUnique
        frag1.extend_from_slice(&(main.len() as u16).to_le_bytes()); // cch
        frag1.push(STR_FLAG_RICH_TEXT | STR_FLAG_EXT); // flags
        frag1.push(0x01); // cRun low byte (cRun=1)

        // Continuation starts with option flags byte, then remaining string bytes.
        let mut frag2 = Vec::new();
        frag2.push(0); // continued segment compressed
        frag2.push(0x00); // cRun high byte
        frag2.extend_from_slice(&(ext_payload.len() as u32).to_le_bytes()); // cbExtRst
        frag2.extend_from_slice(main.as_bytes());
        frag2.extend_from_slice(&rg_run);
        frag2.extend_from_slice(&ext_payload);

        let stream = [
            record(RECORD_SST, &frag1),
            record(records::RECORD_CONTINUE, &frag2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let out = parse_biff8_sst_phonetics(&stream, 1252).expect("parse");
        assert_eq!(out, vec![Some(phonetic.into())]);
    }

    #[test]
    fn parses_sst_phonetic_with_ext_payload_split_across_continue_and_preserves_following_string() {
        let main1 = "abc";
        let phonetic = "Z";
        let ext_payload = xl_unicode_string_compressed(phonetic);
        assert_eq!(ext_payload.len(), 4);

        let main2 = "X";

        // Build full SST payload for both strings, then split within the first string's ext bytes.
        let mut payload_full = Vec::new();
        payload_full.extend_from_slice(&2u32.to_le_bytes()); // cstTotal
        payload_full.extend_from_slice(&2u32.to_le_bytes()); // cstUnique

        // String 1: ext-only.
        let mut s1 = Vec::new();
        s1.extend_from_slice(&(main1.len() as u16).to_le_bytes());
        s1.push(STR_FLAG_EXT);
        s1.extend_from_slice(&(ext_payload.len() as u32).to_le_bytes());
        s1.extend_from_slice(main1.as_bytes());
        s1.extend_from_slice(&ext_payload);

        // String 2: simple.
        let s2 = xl_unicode_string_compressed(main2);

        payload_full.extend_from_slice(&s1);
        payload_full.extend_from_slice(&s2);

        // Split after 3 bytes of the ext payload so the last ext byte appears in the CONTINUE.
        let ext_start = 8 /*sst header*/
            + 2 /*cch*/
            + 1 /*flags*/
            + 4 /*cbExtRst*/
            + main1.len();
        let split_at = ext_start + 3;

        let frag1 = payload_full[..split_at].to_vec();
        let remaining = &payload_full[split_at..];

        let mut frag2 = Vec::new();
        frag2.push(0); // continued segment compressed
        frag2.extend_from_slice(remaining);

        let stream = [
            record(RECORD_SST, &frag1),
            record(records::RECORD_CONTINUE, &frag2),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let out = parse_biff8_sst_phonetics(&stream, 1252).expect("parse");
        assert_eq!(out.len(), 2);
        assert_eq!(out[0], Some(phonetic.into()));
        assert_eq!(out[1], None);
    }
}
