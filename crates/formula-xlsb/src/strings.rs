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

pub(crate) fn read_xl_wide_string(
    rr: &mut RecordReader<'_>,
    flags_width: FlagsWidth,
) -> Result<ParsedXlsbString, Error> {
    read_xl_wide_string_impl(rr, flags_width, true)
}

pub(crate) fn read_xl_wide_string_impl(
    rr: &mut RecordReader<'_>,
    flags_width: FlagsWidth,
    preserve_extras: bool,
) -> Result<ParsedXlsbString, Error> {
    Ok(read_xl_wide_string_with_flags(rr, flags_width, preserve_extras)?.1)
}

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
        let parsed = read_xl_wide_string(&mut rr, FlagsWidth::U16).expect("parse string");
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
        let parsed = read_xl_wide_string(&mut rr, FlagsWidth::U16).expect("parse rich string");
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
        let parsed = read_xl_wide_string(&mut rr, FlagsWidth::U16).expect("parse phonetic string");
        assert_eq!(parsed.text, text);
        assert_eq!(parsed.phonetic.as_deref(), Some(phonetic_bytes.as_slice()));

        let next = rr.read_u32().expect("read next field");
        assert_eq!(next, 0x55667788);
    }
}
