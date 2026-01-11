use encoding_rs::{
    Encoding, BIG5, EUC_KR, GBK, SHIFT_JIS, UTF_8, WINDOWS_1250, WINDOWS_1251, WINDOWS_1252,
    WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257, WINDOWS_1258,
    WINDOWS_874,
};

use super::BiffVersion;

pub(crate) fn encoding_for_codepage(codepage: u16) -> &'static Encoding {
    match codepage as u32 {
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
        _ => WINDOWS_1252,
    }
}

pub(crate) fn decode_ansi(bytes: &[u8], encoding: &'static Encoding) -> String {
    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

pub(crate) fn parse_biff_short_string(
    input: &[u8],
    biff: BiffVersion,
    encoding: &'static Encoding,
) -> Result<(String, usize), String> {
    match biff {
        BiffVersion::Biff5 => parse_biff5_short_string(input, encoding),
        BiffVersion::Biff8 => parse_biff8_short_string(input, encoding),
    }
}

/// BIFF5 "short string": 8-bit length prefix followed by ANSI bytes.
pub(crate) fn parse_biff5_short_string(
    input: &[u8],
    encoding: &'static Encoding,
) -> Result<(String, usize), String> {
    let Some((&len, rest)) = input.split_first() else {
        return Err("unexpected end of string".to_string());
    };
    let len = len as usize;
    let bytes = rest
        .get(0..len)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    Ok((decode_ansi(bytes, encoding), 1 + len))
}

/// BIFF8 `ShortXLUnicodeString` [MS-XLS 2.5.293].
pub(crate) fn parse_biff8_short_string(
    input: &[u8],
    encoding: &'static Encoding,
) -> Result<(String, usize), String> {
    if input.len() < 2 {
        return Err("unexpected end of string".to_string());
    }
    let cch = input[0] as usize;
    let flags = input[1];
    let mut offset = 2usize;

    let richtext_runs = if flags & 0x08 != 0 {
        if input.len() < offset + 2 {
            return Err("unexpected end of string".to_string());
        }
        let runs = u16::from_le_bytes([input[offset], input[offset + 1]]) as usize;
        offset += 2;
        runs
    } else {
        0
    };

    let ext_size = if flags & 0x04 != 0 {
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

    let is_unicode = (flags & 0x01) != 0;
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
        decode_ansi(chars, encoding)
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

/// BIFF8 `XLUnicodeString` [MS-XLS 2.5.268] (16-bit length).
pub(crate) fn parse_biff8_unicode_string(
    input: &[u8],
    encoding: &'static Encoding,
) -> Result<(String, usize), String> {
    if input.len() < 3 {
        return Err("unexpected end of string".to_string());
    }

    let cch = u16::from_le_bytes([input[0], input[1]]) as usize;
    let flags = input[2];
    let mut offset = 3usize;

    let richtext_runs = if flags & 0x08 != 0 {
        if input.len() < offset + 2 {
            return Err("unexpected end of string".to_string());
        }
        let runs = u16::from_le_bytes([input[offset], input[offset + 1]]) as usize;
        offset += 2;
        runs
    } else {
        0
    };

    let ext_size = if flags & 0x04 != 0 {
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

    let is_unicode = (flags & 0x01) != 0;
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
        decode_ansi(chars, encoding)
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

pub(crate) fn parse_biff5_short_string_best_effort(
    input: &[u8],
    encoding: &'static Encoding,
) -> Option<String> {
    let (&len, rest) = input.split_first()?;
    let take = (len as usize).min(rest.len());
    Some(decode_ansi(&rest[..take], encoding))
}

pub(crate) fn parse_biff8_unicode_string_best_effort(
    input: &[u8],
    encoding: &'static Encoding,
) -> Option<String> {
    if input.len() < 3 {
        return None;
    }

    let cch = u16::from_le_bytes([input[0], input[1]]) as usize;
    let flags = input[2];
    let mut offset = 3usize;

    if flags & 0x08 != 0 {
        // cRun (optional)
        if input.len() < offset + 2 {
            return Some(String::new());
        }
        offset += 2;
    }

    if flags & 0x04 != 0 {
        // cbExtRst (optional)
        if input.len() < offset + 4 {
            return Some(String::new());
        }
        offset += 4;
    }

    let is_unicode = (flags & 0x01) != 0;
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
        decode_ansi(bytes, encoding)
    })
}

pub(crate) fn parse_biff8_unicode_string_continued(
    fragments: &[&[u8]],
    start_offset: usize,
    encoding: &'static Encoding,
) -> Result<String, String> {
    let mut cursor = FragmentCursor::new(fragments, 0, start_offset);
    cursor.read_biff8_unicode_string(encoding)
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

    fn read_biff8_unicode_string(&mut self, encoding: &'static Encoding) -> Result<String, String> {
        // XLUnicodeString [MS-XLS 2.5.268]
        let cch = self.read_u16_le()? as usize;
        let flags = self.read_u8()?;

        let richtext_runs = if flags & 0x08 != 0 {
            self.read_u16_le()? as usize
        } else {
            0
        };

        let ext_size = if flags & 0x04 != 0 {
            self.read_u32_le()? as usize
        } else {
            0
        };

        let mut is_unicode = (flags & 0x01) != 0;
        let mut remaining_chars = cch;
        let mut out = String::new();

        while remaining_chars > 0 {
            if self.remaining_in_fragment() == 0 {
                // Continuing character bytes into a new CONTINUE fragment: first
                // byte is option flags for the continued segment (fHighByte).
                self.advance_fragment()?;
                let cont_flags = self.read_u8()?;
                is_unicode = (cont_flags & 0x01) != 0;
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
                out.push_str(&decode_ansi(bytes, encoding));
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
        let encoding = encoding_for_codepage(1252);
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
        let out = parse_biff8_unicode_string_continued(&fragments, 0, encoding).expect("parse");
        assert_eq!(out, s);
    }
}

