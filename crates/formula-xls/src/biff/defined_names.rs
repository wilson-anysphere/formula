#![allow(dead_code)]

use super::{records, strings, BiffVersion};

/// BIFF8 `NAME` record id.
///
/// See [MS-XLS] 2.4.150 (NAME).
const RECORD_NAME: u16 = 0x0018;

// BIFF8 string option flags used by `XLUnicodeStringNoCch`.
// See [MS-XLS] 2.5.292.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BiffDefinedName {
    pub(crate) name: String,
    /// Raw BIFF8 `rgce` bytes for the defined name formula.
    pub(crate) rgce: Vec<u8>,
}

pub(crate) fn parse_biff_defined_names(
    workbook_stream: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<Vec<BiffDefinedName>, String> {
    match biff {
        BiffVersion::Biff8 => parse_biff8_defined_names(workbook_stream, codepage),
        // TODO: BIFF5 `NAME` parsing is not yet needed by the importer.
        BiffVersion::Biff5 => Ok(Vec::new()),
    }
}

fn parse_biff8_defined_names(
    workbook_stream: &[u8],
    codepage: u16,
) -> Result<Vec<BiffDefinedName>, String> {
    let mut out = Vec::new();

    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);
    for record in iter {
        let record = match record {
            Ok(record) => record,
            // Best-effort: stop once we hit a malformed record and return what we have.
            Err(_) => break,
        };

        // The `NAME` record lives in the workbook-global substream. Stop if we see the start of the
        // next substream (worksheet BOF), even if the workbook-global EOF is missing.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            RECORD_NAME => {
                if let Ok(name) = parse_biff8_name_record(&record, codepage) {
                    out.push(name);
                }
            }
            records::RECORD_EOF => break,
            _ => {}
        }
    }

    Ok(out)
}

fn allows_continuation(record_id: u16) -> bool {
    record_id == RECORD_NAME
}

fn parse_biff8_name_record(
    record: &records::LogicalBiffRecord<'_>,
    codepage: u16,
) -> Result<BiffDefinedName, String> {
    let fragments: Vec<&[u8]> = record.fragments().collect();
    let mut cursor = FragmentCursor::new(&fragments, 0, 0);

    // Fixed-size `NAME` record header (14 bytes).
    // [MS-XLS] 2.4.150
    let _grbit = cursor.read_u16_le()?;
    let _ch_key = cursor.read_u8()?;
    let cch = cursor.read_u8()? as usize;
    let cce = cursor.read_u16_le()? as usize;
    let _ixals = cursor.read_u16_le()?;
    let _itab = cursor.read_u16_le()?;
    let cch_cust_menu = cursor.read_u8()? as usize;
    let cch_description = cursor.read_u8()? as usize;
    let cch_help_topic = cursor.read_u8()? as usize;
    let cch_status_text = cursor.read_u8()? as usize;

    // `rgchName` (XLUnicodeStringNoCch): flags byte + character bytes.
    let name = cursor.read_biff8_unicode_string_no_cch(cch, codepage)?;

    // `rgce`: parsed formula bytes.
    let rgce = cursor.read_bytes(cce)?;

    // Optional strings (ignored for now, but we need to consume them so continued-string decoding
    // can validate fragment boundaries).
    if cch_cust_menu > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_cust_menu, codepage)?;
    }
    if cch_description > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_description, codepage)?;
    }
    if cch_help_topic > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_help_topic, codepage)?;
    }
    if cch_status_text > 0 {
        let _ = cursor.read_biff8_unicode_string_no_cch(cch_status_text, codepage)?;
    }

    Ok(BiffDefinedName { name, rgce })
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

    fn read_bytes(&mut self, mut n: usize) -> Result<Vec<u8>, String> {
        let mut out = Vec::with_capacity(n);
        while n > 0 {
            let available = self.remaining_in_fragment();
            if available == 0 {
                self.advance_fragment()?;
                continue;
            }
            let take = n.min(available);
            let bytes = self.read_exact_from_current(take)?;
            out.extend_from_slice(bytes);
            n -= take;
        }
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

    fn read_biff8_unicode_string_no_cch(
        &mut self,
        cch: usize,
        codepage: u16,
    ) -> Result<String, String> {
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

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first byte is option
                // flags for the continued segment (fHighByte).
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
        self.skip_bytes(richtext_bytes + ext_size)?;

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

    #[test]
    fn parses_defined_name_with_continued_rgce_bytes() {
        let name = "Name";
        let rgce: Vec<u8> = vec![0x01, 0x02, 0x03, 0x04, 0x05, 0x06];

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        let mut name_str = Vec::new();
        name_str.push(0); // flags (compressed)
        name_str.extend_from_slice(name.as_bytes());

        let first_rgce = &rgce[..2];
        let second_rgce = &rgce[2..];

        let r_bof = record(records::RECORD_BOF_BIFF8, &[0u8; 16]);
        let r_name = record(RECORD_NAME, &[header.clone(), name_str.clone(), first_rgce.to_vec()].concat());
        let r_continue = record(records::RECORD_CONTINUE, second_rgce);
        let r_eof = record(records::RECORD_EOF, &[]);
        let stream = [r_bof, r_name, r_continue, r_eof].concat();

        let names =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252).expect("parse names");
        assert_eq!(
            names,
            vec![BiffDefinedName {
                name: name.to_string(),
                rgce
            }]
        );
    }

    #[test]
    fn parses_defined_name_with_continued_name_string() {
        let name = "ABCDE";
        let rgce: Vec<u8> = vec![0x11, 0x22];

        let mut header = Vec::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        // Split the name string across records after 2 characters.
        let mut first = Vec::new();
        first.extend_from_slice(&header);
        first.push(0); // string flags (compressed)
        first.extend_from_slice(&name.as_bytes()[..2]); // "AB"

        let mut second = Vec::new();
        second.push(0); // continued segment option flags (fHighByte=0)
        second.extend_from_slice(&name.as_bytes()[2..]); // "CDE"
        second.extend_from_slice(&rgce);

        let stream = [
            record(records::RECORD_BOF_BIFF8, &[0u8; 16]),
            record(RECORD_NAME, &first),
            record(records::RECORD_CONTINUE, &second),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let names =
            parse_biff_defined_names(&stream, BiffVersion::Biff8, 1252).expect("parse names");
        assert_eq!(
            names,
            vec![BiffDefinedName {
                name: name.to_string(),
                rgce
            }]
        );
    }
}
