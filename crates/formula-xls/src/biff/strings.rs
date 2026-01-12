use std::collections::BTreeSet;
use std::sync::{Mutex, OnceLock};

use encoding_rs::{
    Encoding, BIG5, EUC_KR, GBK, SHIFT_JIS, UTF_8, WINDOWS_1250, WINDOWS_1251, WINDOWS_1252,
    WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257, WINDOWS_1258,
    WINDOWS_874,
};

use super::BiffVersion;

// BIFF8 string option flags used by ShortXLUnicodeString and XLUnicodeString.
// See [MS-XLS] 2.5.293 and 2.5.268.
const STR_FLAG_HIGH_BYTE: u8 = 0x01;
const STR_FLAG_EXT: u8 = 0x04;
const STR_FLAG_RICH_TEXT: u8 = 0x08;

pub(crate) fn encoding_for_codepage(codepage: u16) -> Option<&'static Encoding> {
    Some(match codepage as u32 {
        874 => WINDOWS_874,
        932 => SHIFT_JIS,
        936 => GBK,
        949 => EUC_KR,
        950 => BIG5,
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1252 => WINDOWS_1252,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        65001 => UTF_8,
        _ => return None,
    })
}

pub(crate) fn decode_ansi(codepage: u16, bytes: &[u8]) -> String {
    if let Some(encoding) = encoding_for_codepage(codepage) {
        let (cow, _, _) = encoding.decode(bytes);
        return cow.into_owned();
    }

    warn_unsupported_codepage(codepage);

    // Lossless byte-to-Unicode mapping (ISO-8859-1-ish): preserve the original BIFF payload and
    // keep ASCII intact even when the codepage isn't supported by `encoding_rs`.
    bytes.iter().copied().map(char::from).collect()
}

fn warn_unsupported_codepage(codepage: u16) {
    static WARNED: OnceLock<Mutex<BTreeSet<u16>>> = OnceLock::new();

    let warned = WARNED.get_or_init(|| Mutex::new(BTreeSet::new()));
    let mut warned = match warned.lock() {
        Ok(guard) => guard,
        Err(poisoned) => poisoned.into_inner(),
    };

    if warned.insert(codepage) {
        log::warn!(
            "unsupported BIFF CODEPAGE {codepage}; decoding 8-bit strings using lossless byte-to-Unicode mapping"
        );
    }
}

pub(crate) fn parse_biff_short_string(
    input: &[u8],
    biff: BiffVersion,
    codepage: u16,
) -> Result<(String, usize), String> {
    match biff {
        BiffVersion::Biff5 => parse_biff5_short_string(input, codepage),
        BiffVersion::Biff8 => parse_biff8_short_string(input, codepage),
    }
}

/// BIFF5 "short string": 8-bit length prefix followed by ANSI bytes.
pub(crate) fn parse_biff5_short_string(
    input: &[u8],
    codepage: u16,
) -> Result<(String, usize), String> {
    let Some((&len, rest)) = input.split_first() else {
        return Err("unexpected end of string".to_string());
    };
    let len = len as usize;
    let bytes = rest
        .get(0..len)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    Ok((decode_ansi(codepage, bytes), 1 + len))
}

/// BIFF8 `ShortXLUnicodeString` [MS-XLS 2.5.293].
pub(crate) fn parse_biff8_short_string(
    input: &[u8],
    codepage: u16,
) -> Result<(String, usize), String> {
    if input.len() < 2 {
        return Err("unexpected end of string".to_string());
    }
    let cch = input[0] as usize;
    let flags = input[1];
    parse_biff8_string_payload(input, cch, flags, 2, codepage)
}

/// BIFF8 `XLUnicodeString` [MS-XLS 2.5.268] (16-bit length).
pub(crate) fn parse_biff8_unicode_string(
    input: &[u8],
    codepage: u16,
) -> Result<(String, usize), String> {
    if input.len() < 3 {
        return Err("unexpected end of string".to_string());
    }

    let cch = u16::from_le_bytes([input[0], input[1]]) as usize;
    let flags = input[2];
    parse_biff8_string_payload(input, cch, flags, 3, codepage)
}

fn parse_biff8_string_payload(
    input: &[u8],
    cch: usize,
    flags: u8,
    mut offset: usize,
    codepage: u16,
) -> Result<(String, usize), String> {
    let richtext_runs = if flags & STR_FLAG_RICH_TEXT != 0 {
        if input.len() < offset + 2 {
            return Err("unexpected end of string".to_string());
        }
        let runs = u16::from_le_bytes([input[offset], input[offset + 1]]) as usize;
        offset += 2;
        runs
    } else {
        0
    };

    let ext_size = if flags & STR_FLAG_EXT != 0 {
        if input.len() < offset + 4 {
            return Err("unexpected end of string".to_string());
        }
        let size = u32::from_le_bytes([
            input[offset],
            input[offset + 1],
            input[offset + 2],
            input[offset + 3],
        ]) as usize;
        offset += 4;
        size
    } else {
        0
    };

    let is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
    let char_bytes = if is_unicode {
        cch.checked_mul(2)
            .ok_or_else(|| "string length overflow".to_string())?
    } else {
        cch
    };

    let chars = input
        .get(offset..offset + char_bytes)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    offset += char_bytes;

    let value = if is_unicode {
        let mut u16s = Vec::with_capacity(cch);
        for chunk in chars.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        decode_ansi(codepage, chars)
    };

    let richtext_bytes = richtext_runs
        .checked_mul(4)
        .ok_or_else(|| "rich text run count overflow".to_string())?;
    if input.len() < offset + richtext_bytes + ext_size {
        return Err("unexpected end of string".to_string());
    }
    offset += richtext_bytes + ext_size;

    Ok((value, offset))
}

pub(crate) fn parse_biff5_short_string_best_effort(input: &[u8], codepage: u16) -> Option<String> {
    let (&len, rest) = input.split_first()?;
    let take = (len as usize).min(rest.len());
    Some(decode_ansi(codepage, &rest[..take]))
}

pub(crate) fn parse_biff8_unicode_string_best_effort(
    input: &[u8],
    codepage: u16,
) -> Option<String> {
    if input.len() < 3 {
        return None;
    }

    let cch = u16::from_le_bytes([input[0], input[1]]) as usize;
    let flags = input[2];
    let mut offset = 3usize;

    if flags & STR_FLAG_RICH_TEXT != 0 {
        // cRun (optional)
        if input.len() < offset + 2 {
            return Some(String::new());
        }
        offset += 2;
    }

    if flags & STR_FLAG_EXT != 0 {
        // cbExtRst (optional)
        if input.len() < offset + 4 {
            return Some(String::new());
        }
        offset += 4;
    }

    let is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
    let bytes_per_char = if is_unicode { 2 } else { 1 };
    let bytes = input.get(offset..).unwrap_or_default();
    let available_chars = bytes.len() / bytes_per_char;
    let take_chars = cch.min(available_chars);
    let take_bytes = take_chars * bytes_per_char;
    let bytes = &bytes[..take_bytes];

    Some(if is_unicode {
        let mut u16s = Vec::with_capacity(take_chars);
        for chunk in bytes.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        decode_ansi(codepage, bytes)
    })
}

pub(crate) fn parse_biff8_unicode_string_continued(
    fragments: &[&[u8]],
    start_offset: usize,
    codepage: u16,
) -> Result<String, String> {
    let mut cursor = FragmentCursor::new(fragments, 0, start_offset);
    cursor.read_biff8_unicode_string(codepage)
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

    fn read_biff8_unicode_string(&mut self, codepage: u16) -> Result<String, String> {
        // XLUnicodeString [MS-XLS 2.5.268]
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

        let mut is_unicode = (flags & STR_FLAG_HIGH_BYTE) != 0;
        let mut remaining_chars = cch;
        let mut out = String::new();

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
            let bytes = self.read_exact_from_current(take_bytes)?;

            if is_unicode {
                let mut u16s = Vec::with_capacity(take_chars);
                for chunk in bytes.chunks_exact(2) {
                    u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                out.push_str(&String::from_utf16_lossy(&u16s));
            } else {
                out.push_str(&decode_ansi(codepage, bytes));
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

    #[test]
    fn parses_biff8_unicode_string_continued_across_fragments() {
        let s = "ABCDE";

        // First fragment contains XLUnicodeString header + partial character data.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&(s.len() as u16).to_le_bytes());
        frag1.push(0); // flags (compressed)
        frag1.extend_from_slice(&s.as_bytes()[..2]);

        // Continuation fragment begins with option flags byte (fHighByte), then remaining bytes.
        let mut frag2 = Vec::new();
        frag2.push(0); // continued segment compressed
        frag2.extend_from_slice(&s.as_bytes()[2..]);

        let fragments: [&[u8]; 2] = [&frag1, &frag2];
        let out = parse_biff8_unicode_string_continued(&fragments, 0, 1252).expect("parse");
        assert_eq!(out, s);
    }

    #[test]
    fn parses_biff8_unicode_string_continued_across_fragments_unicode() {
        let s = "AB";

        // First fragment contains the header and the first UTF-16LE code unit.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&(s.len() as u16).to_le_bytes());
        frag1.push(STR_FLAG_HIGH_BYTE); // flags (unicode)
        frag1.extend_from_slice(&[b'A', 0x00]);

        // Continuation fragment begins with option flags byte (fHighByte), then remaining UTF-16LE bytes.
        let frag2 = [STR_FLAG_HIGH_BYTE, b'B', 0x00];

        let fragments: [&[u8]; 2] = [&frag1, &frag2];
        let out = parse_biff8_unicode_string_continued(&fragments, 0, 1252).expect("parse");
        assert_eq!(out, s);
    }

    #[test]
    fn continued_unicode_string_errors_on_mid_character_split() {
        // cch=1, unicode.
        let mut frag1 = Vec::new();
        frag1.extend_from_slice(&1u16.to_le_bytes());
        frag1.push(STR_FLAG_HIGH_BYTE); // flags (unicode)
        frag1.push(b'A'); // only 1 byte of the 2-byte code unit

        let frag2 = [STR_FLAG_HIGH_BYTE, 0x00]; // cont_flags + remaining byte

        let fragments: [&[u8]; 2] = [&frag1, &frag2];
        let err = parse_biff8_unicode_string_continued(&fragments, 0, 1252).unwrap_err();
        assert_eq!(err, "string continuation split mid-character");
    }

    #[test]
    fn parses_biff8_short_string_compressed_uses_codepage() {
        // BIFF8 ShortXLUnicodeString with `fHighByte=0` stores 8-bit bytes encoded using the
        // workbook code page (CODEPAGE record). In Windows-1251, 0xC0 is Cyrillic 'А' (U+0410).
        let input = [1u8, 0u8, 0xC0];
        let (s, consumed) = parse_biff8_short_string(&input, 1251).expect("parse");
        assert_eq!(consumed, input.len());
        assert_eq!(s, "А");
    }

    #[test]
    fn parses_biff8_short_string_unicode() {
        // "Hi" as UTF-16LE.
        let input = [2u8, STR_FLAG_HIGH_BYTE, b'H', 0x00, b'i', 0x00];
        let (s, consumed) = parse_biff8_short_string(&input, 1252).expect("parse");
        assert_eq!(consumed, input.len());
        assert_eq!(s, "Hi");
    }

    #[test]
    fn parses_biff8_short_string_with_richtext_and_ext() {
        let mut input = Vec::new();
        // cch=3, flags=richtext+ext (compressed)
        input.extend_from_slice(&[3u8, STR_FLAG_RICH_TEXT | STR_FLAG_EXT]);
        input.extend_from_slice(&1u16.to_le_bytes()); // cRun
        input.extend_from_slice(&2u32.to_le_bytes()); // cbExtRst
        input.extend_from_slice(b"abc"); // char data
        input.extend_from_slice(&[0u8; 4]); // rich text runs payload
        input.extend_from_slice(&[0u8; 2]); // ext payload

        let (s, consumed) = parse_biff8_short_string(&input, 1252).expect("parse");
        assert_eq!(consumed, input.len());
        assert_eq!(s, "abc");
    }

    #[test]
    fn parses_biff8_unicode_string_compressed() {
        let mut input = Vec::new();
        input.extend_from_slice(&5u16.to_le_bytes());
        input.push(0x00); // flags (compressed)
        input.extend_from_slice(b"Hello");
        let (s, consumed) = parse_biff8_unicode_string(&input, 1252).expect("parse");
        assert_eq!(consumed, input.len());
        assert_eq!(s, "Hello");
    }

    #[test]
    fn parses_biff8_unicode_string_unicode() {
        let mut input = Vec::new();
        input.extend_from_slice(&2u16.to_le_bytes());
        input.push(STR_FLAG_HIGH_BYTE); // flags (unicode)
        input.extend_from_slice(&[b'H', 0x00, b'i', 0x00]);
        let (s, consumed) = parse_biff8_unicode_string(&input, 1252).expect("parse");
        assert_eq!(consumed, input.len());
        assert_eq!(s, "Hi");
    }

    #[test]
    fn errors_on_truncated_biff8_unicode_string_data() {
        let mut input = Vec::new();
        input.extend_from_slice(&5u16.to_le_bytes());
        input.push(0x00); // flags (compressed)
        input.extend_from_slice(b"Hel"); // truncated
        let err = parse_biff8_unicode_string(&input, 1252).unwrap_err();
        assert_eq!(err, "unexpected end of string");
    }
}
