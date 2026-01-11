//! Minimal BIFF record parsing helpers used by the legacy `.xls` importer.
//!
//! This module is intentionally best-effort: BIFF is large and this importer only
//! needs a handful of workbook-global and worksheet records.

use std::collections::{BTreeMap, HashMap};
use std::io::{Read, Seek};
use std::path::Path;

use formula_model::{ColProperties, DateSystem, RowProperties};

#[derive(Debug, Default)]
pub(crate) struct SheetRowColProperties {
    pub(crate) rows: BTreeMap<u32, RowProperties>,
    pub(crate) cols: BTreeMap<u32, ColProperties>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BiffVersion {
    Biff5,
    Biff8,
}

/// Read the workbook stream bytes from a compound file.
pub(crate) fn read_workbook_stream_from_xls(path: &Path) -> Result<Vec<u8>, String> {
    let mut comp = cfb::open(path).map_err(|err| err.to_string())?;
    let mut stream = open_xls_workbook_stream(&mut comp)?;

    let mut workbook_stream = Vec::new();
    stream
        .read_to_end(&mut workbook_stream)
        .map_err(|err| err.to_string())?;
    Ok(workbook_stream)
}

pub(crate) fn open_xls_workbook_stream<R: Read + Seek>(
    comp: &mut cfb::CompoundFile<R>,
) -> Result<cfb::Stream<R>, String> {
    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(stream) = comp.open_stream(candidate) {
            return Ok(stream);
        }
    }
    Err("missing workbook stream (expected `Workbook` or `Book`)".to_string())
}

pub(crate) fn detect_biff_version(workbook_stream: &[u8]) -> BiffVersion {
    let Some((record_id, data)) = read_biff_record(workbook_stream, 0) else {
        return BiffVersion::Biff8;
    };

    // BOF record type. Use BIFF8 heuristics compatible with calamine.
    if record_id != 0x0809 && record_id != 0x0009 {
        return BiffVersion::Biff8;
    }

    let Some(biff_version) = data.get(0..2).map(|v| u16::from_le_bytes([v[0], v[1]])) else {
        return BiffVersion::Biff8;
    };

    let dt = data
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .unwrap_or(0);

    match biff_version {
        0x0500 => BiffVersion::Biff5,
        0x0600 => BiffVersion::Biff8,
        0 => {
            if dt == 0x1000 {
                BiffVersion::Biff5
            } else {
                BiffVersion::Biff8
            }
        }
        _ => BiffVersion::Biff8,
    }
}

pub(crate) fn read_row_col_properties_from_xls(
    path: &Path,
) -> Result<HashMap<String, SheetRowColProperties>, String> {
    let workbook_stream = read_workbook_stream_from_xls(path)?;
    let biff = detect_biff_version(&workbook_stream);
    let sheets = parse_biff_bound_sheets(&workbook_stream, biff)?;

    let mut out = HashMap::new();
    for (sheet_name, offset) in sheets {
        if offset >= workbook_stream.len() {
            return Err(format!(
                "sheet `{sheet_name}` has out-of-bounds stream offset {offset}"
            ));
        }
        let props = parse_biff_sheet_row_col_properties(&workbook_stream, offset)?;
        out.insert(sheet_name, props);
    }

    Ok(out)
}

/// Workbook-global BIFF records needed for stable number format and date system import.
#[derive(Debug, Clone)]
pub(crate) struct BiffWorkbookGlobals {
    pub(crate) date_system: DateSystem,
    formats: HashMap<u16, String>,
    xfs: Vec<BiffXf>,
}

impl Default for BiffWorkbookGlobals {
    fn default() -> Self {
        Self {
            date_system: DateSystem::Excel1900,
            formats: HashMap::new(),
            xfs: Vec::new(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BiffXfKind {
    Cell,
    Style,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct BiffXf {
    pub(crate) num_fmt_id: u16,
    pub(crate) kind: Option<BiffXfKind>,
}

impl BiffWorkbookGlobals {
    /// Resolve an Excel number format code string for the given `xf_index`.
    ///
    /// Precedence:
    /// 1. `numFmtId == 0` → `None` ("General")
    /// 2. workbook `FORMAT` record → exact code
    /// 3. `formula_format::builtin_format_code(numFmtId)` → built-in code
    /// 4. otherwise → stable placeholder (`__builtin_numFmtId:{numFmtId}`)
    #[allow(dead_code)]
    pub(crate) fn resolve_number_format_code(&self, xf_index: u32) -> Option<String> {
        let xf = self.xfs.get(xf_index as usize)?;
        let num_fmt_id = xf.num_fmt_id;

        if num_fmt_id == 0 {
            return None;
        }

        if let Some(code) = self.formats.get(&num_fmt_id) {
            return Some(code.clone());
        }

        if let Some(code) = formula_format::builtin_format_code(num_fmt_id) {
            return Some(code.to_string());
        }

        Some(format!("__builtin_numFmtId:{num_fmt_id}"))
    }
}

pub(crate) fn read_workbook_globals_from_xls(path: &Path) -> Result<BiffWorkbookGlobals, String> {
    let workbook_stream = read_workbook_stream_from_xls(path)?;
    let biff = detect_biff_version(&workbook_stream);
    parse_biff_workbook_globals(&workbook_stream, biff)
}

pub(crate) fn parse_biff_workbook_globals(
    workbook_stream: &[u8],
    _biff: BiffVersion,
) -> Result<BiffWorkbookGlobals, String> {
    let mut out = BiffWorkbookGlobals::default();

    let mut offset = 0usize;
    let mut saw_eof = false;
    loop {
        let Some((record_id, data)) = read_biff_record(workbook_stream, offset) else {
            break;
        };
        offset = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
            .ok_or_else(|| "BIFF record offset overflow".to_string())?;

        match record_id {
            // 1904 [MS-XLS 2.4.169]
            0x0022 => {
                if data.len() < 2 {
                    return Err("1904 record too short".to_string());
                }
                let flag = u16::from_le_bytes([data[0], data[1]]);
                if flag != 0 {
                    out.date_system = DateSystem::Excel1904;
                }
            }
            // FORMAT / FORMAT2 [MS-XLS 2.4.90]
            0x041E | 0x001E => {
                let (num_fmt_id, code) = parse_biff_format_record(record_id, data)?;
                out.formats.insert(num_fmt_id, code);
            }
            // XF [MS-XLS 2.4.353]
            0x00E0 => {
                let xf = parse_biff_xf_record(data)?;
                out.xfs.push(xf);
            }
            // EOF terminates the workbook global substream.
            0x000A => {
                saw_eof = true;
                break;
            }
            _ => {}
        }
    }

    if !saw_eof {
        return Err("unexpected end of workbook globals stream (missing EOF)".to_string());
    }

    Ok(out)
}

fn parse_biff_xf_record(data: &[u8]) -> Result<BiffXf, String> {
    if data.len() < 4 {
        return Err("XF record too short".to_string());
    }

    let num_fmt_id = u16::from_le_bytes([data[2], data[3]]);

    // Optional: in BIFF5/8 this is part of the "type/protection" flags field and bit 2 is `fStyle`.
    let kind = data.get(4..6).map(|bytes| {
        let flags = u16::from_le_bytes([bytes[0], bytes[1]]);
        if (flags & 0x0004) != 0 {
            BiffXfKind::Style
        } else {
            BiffXfKind::Cell
        }
    });

    Ok(BiffXf { num_fmt_id, kind })
}

fn parse_biff_format_record(record_id: u16, data: &[u8]) -> Result<(u16, String), String> {
    if data.len() < 2 {
        return Err("FORMAT record too short".to_string());
    }
    let num_fmt_id = u16::from_le_bytes([data[0], data[1]]);
    let rest = &data[2..];

    let (mut code, _) = match record_id {
        // BIFF8 FORMAT uses `XLUnicodeString` (16-bit length).
        0x041E => parse_biff8_unicode_string(rest)?,
        // BIFF5 FORMAT2 uses a short ANSI string (8-bit length).
        0x001E => parse_biff5_short_string(rest)?,
        _ => return Err(format!("unexpected FORMAT record id 0x{record_id:04X}")),
    };

    // Excel stores some strings with embedded NUL bytes; follow BoundSheet parsing and strip them.
    code = code.replace('\0', "");
    Ok((num_fmt_id, code))
}

pub(crate) fn parse_biff_bound_sheets(
    workbook_stream: &[u8],
    biff: BiffVersion,
) -> Result<Vec<(String, usize)>, String> {
    let mut offset = 0usize;
    let mut out = Vec::new();

    loop {
        let Some((record_id, data)) = read_biff_record(workbook_stream, offset) else {
            break;
        };
        offset = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
            .ok_or_else(|| "BIFF record offset overflow".to_string())?;

        match record_id {
            // BoundSheet8 [MS-XLS 2.4.28]
            0x0085 => {
                if data.len() < 7 {
                    return Err("BoundSheet8 record too short".to_string());
                }

                let sheet_offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                let (name, _) = parse_biff_short_string(&data[6..], biff)?;
                let name = name.replace('\0', "");
                out.push((name, sheet_offset));
            }
            // EOF terminates the workbook global substream.
            0x000A => break,
            _ => {}
        }
    }

    Ok(out)
}

pub(crate) fn parse_biff_sheet_row_col_properties(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetRowColProperties, String> {
    let mut props = SheetRowColProperties::default();
    let mut offset = start;

    loop {
        let Some((record_id, data)) = read_biff_record(workbook_stream, offset) else {
            break;
        };
        offset = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
            .ok_or_else(|| "BIFF record offset overflow".to_string())?;

        match record_id {
            // ROW [MS-XLS 2.4.184]
            0x0208 => {
                if data.len() < 16 {
                    return Err("ROW record too short".to_string());
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let height_options = u16::from_le_bytes([data[6], data[7]]);
                let height_twips = height_options & 0x7FFF;
                let default_height = (height_options & 0x8000) != 0;
                let options = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let hidden = (options & 0x0000_0020) != 0;

                let height = (!default_height && height_twips > 0)
                    .then_some(height_twips as f32 / 20.0);

                if hidden || height.is_some() {
                    let entry = props.rows.entry(row).or_default();
                    if let Some(height) = height {
                        entry.height = Some(height);
                    }
                    if hidden {
                        entry.hidden = true;
                    }
                }
            }
            // COLINFO [MS-XLS 2.4.48]
            0x007D => {
                if data.len() < 12 {
                    return Err("COLINFO record too short".to_string());
                }
                let first_col = u16::from_le_bytes([data[0], data[1]]) as u32;
                let last_col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let width_raw = u16::from_le_bytes([data[4], data[5]]);
                let options = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (options & 0x0001) != 0;

                let width = (width_raw > 0).then_some(width_raw as f32 / 256.0);

                if hidden || width.is_some() {
                    for col in first_col..=last_col {
                        let entry = props.cols.entry(col).or_default();
                        if let Some(width) = width {
                            entry.width = Some(width);
                        }
                        if hidden {
                            entry.hidden = true;
                        }
                    }
                }
            }
            // EOF terminates the sheet substream.
            0x000A => break,
            _ => {}
        }
    }

    Ok(props)
}

pub(crate) fn read_biff_record(workbook_stream: &[u8], offset: usize) -> Option<(u16, &[u8])> {
    let header = workbook_stream.get(offset..offset + 4)?;
    let record_id = u16::from_le_bytes([header[0], header[1]]);
    let len = u16::from_le_bytes([header[2], header[3]]) as usize;
    let data_start = offset + 4;
    let data_end = data_start.checked_add(len)?;
    let data = workbook_stream.get(data_start..data_end)?;
    Some((record_id, data))
}

fn parse_biff_short_string(input: &[u8], biff: BiffVersion) -> Result<(String, usize), String> {
    match biff {
        BiffVersion::Biff5 => parse_biff5_short_string(input),
        BiffVersion::Biff8 => parse_biff8_short_string(input),
    }
}

fn parse_biff5_short_string(input: &[u8]) -> Result<(String, usize), String> {
    let Some((&len, rest)) = input.split_first() else {
        return Err("unexpected end of string".to_string());
    };
    let len = len as usize;
    let bytes = rest
        .get(0..len)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    Ok((String::from_utf8_lossy(bytes).into_owned(), 1 + len))
}

fn parse_biff8_short_string(input: &[u8]) -> Result<(String, usize), String> {
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

    let name = if is_unicode {
        let mut u16s = Vec::with_capacity(cch);
        for chunk in chars.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        String::from_utf8_lossy(chars).into_owned()
    };

    let richtext_bytes = richtext_runs
        .checked_mul(4)
        .ok_or_else(|| "rich text run count overflow".to_string())?;
    if input.len() < offset + richtext_bytes + ext_size {
        return Err("unexpected end of string".to_string());
    }
    offset += richtext_bytes + ext_size;

    Ok((name, offset))
}

fn parse_biff8_unicode_string(input: &[u8]) -> Result<(String, usize), String> {
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

    let s = if is_unicode {
        let mut u16s = Vec::with_capacity(cch);
        for chunk in chars.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        String::from_utf8_lossy(chars).into_owned()
    };

    let richtext_bytes = richtext_runs
        .checked_mul(4)
        .ok_or_else(|| "rich text run count overflow".to_string())?;
    if input.len() < offset + richtext_bytes + ext_size {
        return Err("unexpected end of string".to_string());
    }
    offset += richtext_bytes + ext_size;

    Ok((s, offset))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u16, data: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + data.len());
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
        out
    }

    #[test]
    fn parses_globals_date_system_formats_and_xfs_biff8() {
        // 1904 record payload: f1904 = 1.
        let r_1904 = record(0x0022, &[1, 0]);

        // FORMAT record: id=164, code="0.00" as XLUnicodeString (compressed).
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&164u16.to_le_bytes());
        fmt_payload.extend_from_slice(&4u16.to_le_bytes()); // cch
        fmt_payload.push(0); // flags (compressed)
        fmt_payload.extend_from_slice(b"0.00");
        let r_fmt = record(0x041E, &fmt_payload);

        // XF record referencing numFmtId=164, cell xf (fStyle=0).
        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&164u16.to_le_bytes());
        xf_payload[4..6].copy_from_slice(&0u16.to_le_bytes());
        let r_xf = record(0x00E0, &xf_payload);

        let r_eof = record(0x000A, &[]);

        let mut stream = Vec::new();
        stream.extend_from_slice(&r_1904);
        stream.extend_from_slice(&r_fmt);
        stream.extend_from_slice(&r_xf);
        stream.extend_from_slice(&r_eof);

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1904);
        assert_eq!(globals.resolve_number_format_code(0).as_deref(), Some("0.00"));
    }

    #[test]
    fn resolves_builtins_and_placeholders() {
        let r_1900 = record(0x0022, &[0, 0]);

        // Two XF records: one built-in (14), one unknown (60), and one General (0).
        let mut xf14 = vec![0u8; 20];
        xf14[2..4].copy_from_slice(&14u16.to_le_bytes());
        let mut xf60 = vec![0u8; 20];
        xf60[2..4].copy_from_slice(&60u16.to_le_bytes());
        let mut xf0 = vec![0u8; 20];
        xf0[2..4].copy_from_slice(&0u16.to_le_bytes());

        let stream = [
            r_1900,
            record(0x00E0, &xf14),
            record(0x00E0, &xf60),
            record(0x00E0, &xf0),
            record(0x000A, &[]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("m/d/yyyy")
        );
        assert_eq!(
            globals.resolve_number_format_code(1).as_deref(),
            Some("__builtin_numFmtId:60")
        );
        assert_eq!(globals.resolve_number_format_code(2), None);
        assert_eq!(globals.resolve_number_format_code(99), None);
    }

    #[test]
    fn parses_biff5_format_strings_and_strips_nuls() {
        // FORMAT2 record: id=200, "0\\0.00" (embedded NUL) as short ANSI string.
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&200u16.to_le_bytes());
        fmt_payload.push(5); // cch (including NUL)
        fmt_payload.extend_from_slice(b"0\0.00");
        let r_fmt = record(0x001E, &fmt_payload);

        let mut xf_payload = vec![0u8; 16];
        xf_payload[2..4].copy_from_slice(&200u16.to_le_bytes());

        let stream = [
            r_fmt,
            record(0x00E0, &xf_payload),
            record(0x000A, &[]),
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff5).expect("parse");
        assert_eq!(globals.resolve_number_format_code(0).as_deref(), Some("0.00"));
    }
}
