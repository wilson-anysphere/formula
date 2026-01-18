use std::collections::HashMap;
use std::io::Cursor;

use crate::parser::{Biff12Reader, Error};
use formula_format::Locale;

/// Resolved style information for a single XF record referenced by cell `style` indices.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StyleInfo {
    /// Number format id (`numFmtId` / `ifmt`) as stored in the XF record.
    pub num_fmt_id: u16,
    /// Resolved number format code string (built-in or custom).
    pub number_format: Option<String>,
    /// Whether the number format looks like a date/time.
    pub is_date_time: bool,
}

/// Workbook style table derived from `xl/styles.bin`.
///
/// This is intentionally minimal: we currently only extract enough information to
/// resolve XF indices to number format strings and identify date/time formats.
#[derive(Debug, Clone, Default)]
pub struct Styles {
    xfs: Vec<StyleInfo>,
}

impl Styles {
    /// Parse `xl/styles.bin` and return a best-effort [`Styles`] mapping.
    ///
    /// The parser ignores record types it doesn't understand and only extracts:
    /// - custom number formats (`BrtFmt` within the `BrtBeginFmts` section)
    /// - XF records in the `BrtBeginCellXfs` section (the indices referenced by cells)
    pub fn parse(styles_bin: &[u8]) -> Result<Self, Error> {
        Self::parse_with_locale(styles_bin, Locale::en_us())
    }

    /// Parse `xl/styles.bin` using locale-aware built-in number formats.
    ///
    /// XLSB/XLSX styles often reference built-in number formats by id without an
    /// explicit format code. Excel's built-in table is locale-dependent, so
    /// callers that know the workbook locale can pass it here to resolve
    /// built-ins more accurately.
    pub fn parse_with_locale(styles_bin: &[u8], locale: Locale) -> Result<Self, Error> {
        // Record ids (BIFF12 / MS-XLSB).
        const BEGIN_FMTS: u32 = 0x0118;
        const END_FMTS: u32 = 0x0119;
        const BEGIN_CELL_XFS: u32 = 0x0122;
        const END_CELL_XFS: u32 = 0x0123;

        let mut reader = Biff12Reader::new(Cursor::new(styles_bin));
        let mut buf = Vec::new();

        let mut custom_fmts: HashMap<u16, String> = HashMap::new();
        let mut xf_num_fmts: Vec<u16> = Vec::new();

        let mut in_fmts = false;
        let mut in_cell_xfs = false;

        while let Some(rec) = reader.read_record(&mut buf)? {
            match rec.id {
                BEGIN_FMTS => in_fmts = true,
                END_FMTS => in_fmts = false,
                BEGIN_CELL_XFS => in_cell_xfs = true,
                END_CELL_XFS => in_cell_xfs = false,
                _ => {
                    if in_fmts {
                        if let Some((id, code)) = parse_fmt_record(rec.data) {
                            custom_fmts.insert(id, code);
                        }
                    } else if in_cell_xfs {
                        if let Some(num_fmt_id) = parse_xf_record_num_fmt_id(rec.data) {
                            xf_num_fmts.push(num_fmt_id);
                        }
                    }
                }
            }
        }

        let xfs = xf_num_fmts
            .into_iter()
            .map(|num_fmt_id| {
                let number_format = custom_fmts.get(&num_fmt_id).cloned().or_else(|| {
                    formula_format::builtin_format_code_with_locale(num_fmt_id, locale)
                        .map(|s| s.into_owned())
                }).or_else(|| {
                    // Excel reserves many built-in ids beyond 0–49. XLSB/XLSX files can reference
                    // those ids without providing an explicit format code. Preserve the id for
                    // round-trip even if we don't know the code yet.
                    if num_fmt_id < 164 {
                        Some(format!(
                            "{}{num_fmt_id}",
                            formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX
                        ))
                    } else {
                        None
                    }
                });

                // If we have a format string, run a light heuristic. If we don't, fall back
                // to common built-in date/time ids (Excel reserves many date formats).
                let is_date_time = match number_format.as_deref() {
                    // `__builtin_numFmtId:<id>` is an internal placeholder, not a real format code.
                    // Use the numeric id range heuristic instead of scanning it as text.
                    Some(fmt)
                        if fmt.starts_with(formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX) =>
                    {
                        is_reserved_datetime_format_id(num_fmt_id)
                    }
                    Some(fmt) => looks_like_datetime(fmt),
                    None => is_reserved_datetime_format_id(num_fmt_id),
                };

                StyleInfo {
                    num_fmt_id,
                    number_format,
                    is_date_time,
                }
            })
            .collect();

        Ok(Self { xfs })
    }

    /// Returns the resolved style info for a cell `style` id (XF index).
    pub fn get(&self, style_id: u32) -> Option<&StyleInfo> {
        self.xfs.get(style_id as usize)
    }

    /// Number of parsed XF records.
    pub fn len(&self) -> usize {
        self.xfs.len()
    }

    pub fn is_empty(&self) -> bool {
        self.xfs.is_empty()
    }
}

fn parse_xf_record_num_fmt_id(data: &[u8]) -> Option<u16> {
    // BrtXF is a fairly large record, but the number format id is stored early.
    // For now, treat the first u16 as numFmtId (best-effort).
    let bytes: [u8; 2] = data.get(0..2)?.try_into().ok()?;
    Some(u16::from_le_bytes(bytes))
}

fn parse_fmt_record(data: &[u8]) -> Option<(u16, String)> {
    parse_fmt_record_u16(data).or_else(|| parse_fmt_record_u32(data))
}

fn parse_fmt_record_u16(data: &[u8]) -> Option<(u16, String)> {
    let mut rr = RecordReader::new(data);
    let id = rr.read_u16().ok()?;
    let code = rr.read_utf16_string().ok()?;
    Some((id, code))
}

fn parse_fmt_record_u32(data: &[u8]) -> Option<(u16, String)> {
    let mut rr = RecordReader::new(data);
    let id = rr.read_u32().ok()?;
    let id = u16::try_from(id).ok()?;
    let code = rr.read_utf16_string().ok()?;
    Some((id, code))
}

fn is_reserved_datetime_format_id(num_fmt_id: u16) -> bool {
    // Standard built-in date/time formats in OOXML are 14-22 and 45-47. Excel also
    // uses additional reserved ids for locale-specific date/time formats.
    matches!(num_fmt_id, 14..=22 | 45..=47 | 27..=36 | 50..=58)
}

fn looks_like_datetime(section: &str) -> bool {
    // Based on the heuristic in `formula-format` (kept local to avoid exposing
    // `formula_format::datetime::looks_like_datetime`).
    let mut in_quotes = false;
    let mut prev: Option<char> = None;
    let mut chars = section.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => {
                // Skip escaped char.
                let _ = chars.next();
            }
            '[' => {
                // Elapsed time: [h], [m], [s]
                let mut content = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    content.push(c);
                }
                if content.eq_ignore_ascii_case("h")
                    || content.eq_ignore_ascii_case("hh")
                    || content.eq_ignore_ascii_case("m")
                    || content.eq_ignore_ascii_case("mm")
                    || content.eq_ignore_ascii_case("s")
                    || content.eq_ignore_ascii_case("ss")
                {
                    return true;
                }
            }
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => return true,
            'm' | 'M' => {
                // `m` is ambiguous; treat it as datetime only if it is likely
                // part of a date/time expression (e.g. next to y/d/h/s).
                if let Some(p) = prev {
                    if matches!(p, 'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S') {
                        return true;
                    }
                }
                if let Some(n) = chars.peek().copied() {
                    if matches!(n, 'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S') {
                        return true;
                    }
                }
            }
            'a' | 'A' => {
                // AM/PM or A/P markers (case-insensitive).
                let mut probe = String::new();
                probe.push(ch);
                let mut clone = chars.clone();
                for _ in 0..4 {
                    if let Some(c) = clone.next() {
                        probe.push(c);
                    } else {
                        break;
                    }
                }
                if probe
                    .get(.."am/pm".len())
                    .is_some_and(|p| p.eq_ignore_ascii_case("am/pm"))
                    || probe
                        .get(.."a/p".len())
                        .is_some_and(|p| p.eq_ignore_ascii_case("a/p"))
                {
                    return true;
                }
            }
            _ => {}
        }

        prev = Some(ch);
    }

    false
}

struct RecordReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> RecordReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn read_u16(&mut self) -> Result<u16, Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + 2)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += 2;
        Ok(u16::from_le_bytes([raw[0], raw[1]]))
    }

    fn read_u32(&mut self) -> Result<u32, Error> {
        let raw = self
            .data
            .get(self.offset..self.offset + 4)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += 4;
        Ok(u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]))
    }

    fn read_utf16_string(&mut self) -> Result<String, Error> {
        let len_chars = self.read_u32()? as usize;
        let byte_len = len_chars.checked_mul(2).ok_or(Error::UnexpectedEof)?;
        let raw = self
            .data
            .get(self.offset..self.offset + byte_len)
            .ok_or(Error::UnexpectedEof)?;
        self.offset += byte_len;

        // Avoid allocating an intermediate `Vec<u16>` for attacker-controlled string lengths;
        // decode UTF-16LE directly into a `String`.
        let mut out = String::new();
        let _ = out.try_reserve(raw.len());
        let iter = raw
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
        for decoded in std::char::decode_utf16(iter) {
            match decoded {
                Ok(ch) => out.push(ch),
                Err(_) => out.push('\u{FFFD}'),
            }
        }
        Ok(out)
    }
}
