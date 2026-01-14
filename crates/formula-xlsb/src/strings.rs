use crate::parser::{Error, RecordReader};

/// Opaque rich-text payload for BIFF12 "wide strings".
///
/// XLSB stores rich text as an array of formatting runs following the UTF-16 text.
/// We currently preserve the run bytes opaquely for round-trip.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueRichText {
    /// Raw `StrRun` bytes (run count is stored separately in the record).
    pub runs: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedXlsbString {
    pub text: String,
    pub rich: Option<OpaqueRichText>,
    /// Opaque phonetic / extended string bytes for round-trip.
    pub phonetic: Option<Vec<u8>>,
}

impl ParsedXlsbString {
    /// Best-effort decode of the phonetic guide (furigana) text from the opaque phonetic block.
    ///
    /// XLSB stores an "extended string"/phonetic payload after the main UTF-16 text when the
    /// `FLAG_PHONETIC` bit is set. The raw bytes are length-delimited by `cb` in the surrounding
    /// record and preserved in [`ParsedXlsbString::phonetic`] for round-trip.
    ///
    /// The exact binary layout of the phonetic block is not yet implemented here. Empirically it
    /// appears to embed a length-prefixed UTF-16LE string containing the visible phonetic text,
    /// along with additional metadata (e.g. run mappings).
    ///
    /// This function is intentionally tolerant:
    /// - If the payload is missing or cannot be decoded, it returns `None`.
    /// - It does **not** error the workbook parse/export path.
    ///
    /// Assumed layout (best-effort):
    ///
    /// ```text
    /// [.. header ..][cch: u16|u32][phonetic_text: UTF-16LE (cch code units)] [.. trailing ..]
    /// ```
    ///
    /// TODO: Replace the heuristic with a spec-backed parser once the relevant MS-XLSB section is
    /// identified.
    pub fn phonetic_text(&self) -> Option<String> {
        let data = self.phonetic.as_deref()?;
        parse_phonetic_text_best_effort(data)
    }
}

fn parse_phonetic_text_best_effort(data: &[u8]) -> Option<String> {
    // Try to locate a length-prefixed UTF-16LE string inside the phonetic block.
    //
    // We scan for common patterns (u16/u32 character counts) near the start of the block.
    // This is deliberately conservative to avoid decoding arbitrary binary data.
    let mut best: Option<(i32, usize, String)> = None;

    // Scan a limited window to keep this bounded even if `cb` is large.
    // The phonetic string header is typically near the beginning.
    let scan_len = data.len().min(128);
    for offset in 0..scan_len {
        // Prefer aligned offsets (UTF-16 structures are usually 2-byte aligned), but still accept
        // odd offsets for robustness.
        if let Some(s) = decode_len_prefixed_utf16le_u16(data, offset) {
            if let Some(score) = score_candidate(&s) {
                best = pick_better(best, score, offset, s);
            }
        }
        if let Some(s) = decode_len_prefixed_utf16le_u32(data, offset) {
            if let Some(score) = score_candidate(&s) {
                best = pick_better(best, score, offset, s);
            }
        }
    }

    best.map(|(_, _, s)| s)
}

fn pick_better(
    best: Option<(i32, usize, String)>,
    score: i32,
    offset: usize,
    s: String,
) -> Option<(i32, usize, String)> {
    match best {
        None => Some((score, offset, s)),
        Some((best_score, best_offset, best_s)) => {
            // Prefer higher score; if tied, prefer longer strings; if still tied, prefer earlier
            // offsets.
            let best_len = best_s.chars().count();
            let len = s.chars().count();
            if score > best_score
                || (score == best_score && len > best_len)
                || (score == best_score && len == best_len && offset < best_offset)
            {
                Some((score, offset, s))
            } else {
                Some((best_score, best_offset, best_s))
            }
        }
    }
}

fn score_candidate(s: &str) -> Option<i32> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }

    let mut good = 0i32;
    let mut bad = 0i32;
    for ch in s.chars() {
        match ch {
            // Replacement char produced by lossy UTF-16 decoding; treat as suspicious.
            '\u{FFFD}' => bad += 3,
            // NUL is unlikely in visible phonetic text.
            '\u{0}' => bad += 3,
            _ if ch.is_control() => bad += 1,
            _ => good += 1,
        }
    }

    if good == 0 {
        return None;
    }

    // Penalize candidates with lots of "bad" chars.
    let score = good.saturating_mul(2).saturating_sub(bad);
    Some(score)
}

fn decode_len_prefixed_utf16le_u16(data: &[u8], offset: usize) -> Option<String> {
    let raw_len: [u8; 2] = data.get(offset..offset + 2)?.try_into().ok()?;
    let cch = u16::from_le_bytes(raw_len) as usize;
    if cch == 0 {
        return None;
    }
    let bytes_needed = cch.checked_mul(2)?;
    let start = offset.checked_add(2)?;
    let end = start.checked_add(bytes_needed)?;
    let raw = data.get(start..end)?;
    decode_utf16le_lossy(raw)
}

fn decode_len_prefixed_utf16le_u32(data: &[u8], offset: usize) -> Option<String> {
    let raw_len: [u8; 4] = data.get(offset..offset + 4)?.try_into().ok()?;
    let cch = u32::from_le_bytes(raw_len) as usize;
    // Guard against absurd lengths from random data.
    if cch == 0 || cch > 1_000_000 {
        return None;
    }
    let bytes_needed = cch.checked_mul(2)?;
    let start = offset.checked_add(4)?;
    let end = start.checked_add(bytes_needed)?;
    let raw = data.get(start..end)?;
    decode_utf16le_lossy(raw)
}

fn decode_utf16le_lossy(raw: &[u8]) -> Option<String> {
    if raw.len() % 2 != 0 {
        return None;
    }
    // Avoid allocating an intermediate `Vec<u16>` for attacker-controlled inputs; decode
    // directly into a `String`.
    let mut out = String::with_capacity(raw.len());
    let iter = raw
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    for decoded in std::char::decode_utf16(iter) {
        match decoded {
            Ok(ch) => out.push(ch),
            Err(_) => out.push('\u{FFFD}'),
        }
    }
    Some(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FlagsWidth {
    U8,
    #[allow(dead_code)]
    U16,
}

// BIFF12 "wide string" flags.
//
// MS-XLSB stores rich text runs and phonetic/extended data after the main UTF-16 text.
// These bits mirror what other XLSB readers (e.g. pyxlsb) use.
const FLAG_RICH: u16 = 0x0001;
const FLAG_PHONETIC: u16 = 0x0002;

// Size (in bytes) of a single rich text formatting run.
//
// MS-XLSB `StrRun` entries are 8 bytes:
//   [ich: u32][ifnt: u16][reserved: u16]
// (see also the shared string parser in `parser.rs`).
const RICH_RUN_BYTE_LEN: usize = 8;

pub(crate) fn read_xl_wide_string_with_flags(
    rr: &mut RecordReader<'_>,
    flags_width: FlagsWidth,
    preserve_extras: bool,
) -> Result<(u16, ParsedXlsbString), Error> {
    let cch = rr.read_u32()? as usize;
    let flags = match flags_width {
        FlagsWidth::U8 => rr.read_u8()? as u16,
        FlagsWidth::U16 => rr.read_u16()?,
    };

    let text = rr.read_utf16_chars(cch)?;

    let mut rich = None;
    let mut phonetic = None;

    if flags & FLAG_RICH != 0 {
        let c_run = rr.read_u32()? as usize;
        let run_bytes = c_run
            .checked_mul(RICH_RUN_BYTE_LEN)
            .ok_or(Error::UnexpectedEof)?;
        if preserve_extras {
            rich = Some(OpaqueRichText {
                runs: rr.read_slice(run_bytes)?.to_vec(),
            });
        } else {
            rr.skip(run_bytes)?;
        }
    }

    if flags & FLAG_PHONETIC != 0 {
        let cb = rr.read_u32()? as usize;
        if preserve_extras {
            phonetic = Some(rr.read_slice(cb)?.to_vec());
        } else {
            rr.skip(cb)?;
        }
    }

    Ok((
        flags,
        ParsedXlsbString {
            text,
            rich,
            phonetic,
        },
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode_utf16(s: &str) -> Vec<u8> {
        s.encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect::<Vec<u8>>()
    }

    #[test]
    fn phonetic_text_decodes_utf16_from_synthetic_payload() {
        // Synthetic layout:
        //   [reserved: u16][cch: u16][utf16le chars...][trailing bytes...]
        //
        // This matches the heuristic parser expectations without requiring a full MS-XLSB spec
        // implementation.
        let phonetic = "フリガナ";
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u16.to_le_bytes()); // reserved/marker
        payload.extend_from_slice(&(phonetic.encode_utf16().count() as u16).to_le_bytes());
        payload.extend_from_slice(&encode_utf16(phonetic));
        payload.extend_from_slice(&[0xAA, 0xBB, 0xCC]); // trailing bytes (ignored)

        let s = ParsedXlsbString {
            text: "Base".to_string(),
            rich: None,
            phonetic: Some(payload),
        };

        assert_eq!(s.phonetic_text(), Some(phonetic.to_string()));
    }

    #[test]
    fn parses_wide_string_and_leaves_offset_for_following_fields() {
        // [cch:u32][flags:u16][chars...][cce:u32]
        let text = "He said \"Hi\"";
        let chars = encode_utf16(text);
        let cch = (chars.len() / 2) as u32;

        let mut data = Vec::new();
        data.extend_from_slice(&cch.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes()); // flags
        data.extend_from_slice(&chars);
        data.extend_from_slice(&0xAABBCCDDu32.to_le_bytes()); // sentinel "next field"

        let mut rr = RecordReader::new(&data);
        let parsed =
            read_xl_wide_string_with_flags(&mut rr, FlagsWidth::U16, true).expect("parse string").1;
        assert_eq!(parsed.text, text);

        let next = rr.read_u32().expect("read next field");
        assert_eq!(next, 0xAABBCCDD);
    }

    #[test]
    fn skips_rich_text_runs_and_aligns_to_next_field() {
        // Rich string with one dummy run (4 bytes).
        let text = "Rich";
        let chars = encode_utf16(text);
        let cch = (chars.len() / 2) as u32;

        let mut data = Vec::new();
        data.extend_from_slice(&cch.to_le_bytes());
        data.extend_from_slice(&(FLAG_RICH as u16).to_le_bytes());
        data.extend_from_slice(&chars);
        data.extend_from_slice(&1u32.to_le_bytes()); // cRun
        data.extend_from_slice(&[0u8; RICH_RUN_BYTE_LEN]); // rgRun
        data.extend_from_slice(&0x11223344u32.to_le_bytes());

        let mut rr = RecordReader::new(&data);
        let parsed =
            read_xl_wide_string_with_flags(&mut rr, FlagsWidth::U16, true).expect("parse rich string").1;
        assert_eq!(parsed.text, text);
        assert!(parsed.rich.is_some());

        let next = rr.read_u32().expect("read next field");
        assert_eq!(next, 0x11223344);
    }

    #[test]
    fn skips_phonetic_block_and_aligns_to_next_field() {
        let text = "Pho";
        let chars = encode_utf16(text);
        let cch = (chars.len() / 2) as u32;

        let phonetic_bytes = vec![1u8, 2, 3, 4, 5];

        let mut data = Vec::new();
        data.extend_from_slice(&cch.to_le_bytes());
        data.extend_from_slice(&(FLAG_PHONETIC as u16).to_le_bytes());
        data.extend_from_slice(&chars);
        data.extend_from_slice(&(phonetic_bytes.len() as u32).to_le_bytes());
        data.extend_from_slice(&phonetic_bytes);
        data.extend_from_slice(&0x55667788u32.to_le_bytes());

        let mut rr = RecordReader::new(&data);
        let parsed = read_xl_wide_string_with_flags(&mut rr, FlagsWidth::U16, true)
            .expect("parse phonetic string")
            .1;
        assert_eq!(parsed.text, text);
        assert_eq!(parsed.phonetic.as_deref(), Some(phonetic_bytes.as_slice()));

        let next = rr.read_u32().expect("read next field");
        assert_eq!(next, 0x55667788);
    }
}
