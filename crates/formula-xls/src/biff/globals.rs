use std::collections::HashMap;

use encoding_rs::Encoding;
use formula_model::DateSystem;

use super::{records, strings, BiffVersion};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct BoundSheetInfo {
    pub(crate) name: String,
    pub(crate) offset: usize,
}

fn biff_codepage(workbook_stream: &[u8]) -> u16 {
    let mut iter = match records::BiffRecordIter::from_offset(workbook_stream, 0) {
        Ok(iter) => iter,
        Err(_) => return 1252,
    };

    while let Some(record) = iter.next() {
        let record = match record {
            Ok(record) => record,
            Err(_) => break,
        };

        // BOF indicates the start of a new substream; the workbook globals
        // contain a single BOF at offset 0, so a second BOF means we're past
        // the globals section.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            // CODEPAGE [MS-XLS 2.4.52]
            0x0042 => {
                if record.data.len() >= 2 {
                    return u16::from_le_bytes([record.data[0], record.data[1]]);
                }
            }
            // EOF terminates the workbook global substream.
            0x000A => break,
            _ => {}
        }
    }

    // Default "ANSI" codepage used by Excel on Windows.
    1252
}

pub(crate) fn parse_biff_bound_sheets(
    workbook_stream: &[u8],
    biff: BiffVersion,
) -> Result<Vec<BoundSheetInfo>, String> {
    let encoding = strings::encoding_for_codepage(biff_codepage(workbook_stream));
    let mut out = Vec::new();

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, 0)?;
    while let Some(record) = iter.next() {
        let record = record?;

        // Same rationale as `parse_biff_workbook_globals`: stop once we reach the
        // BOF record for the next substream.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }

        match record.record_id {
            // BoundSheet8 [MS-XLS 2.4.28]
            0x0085 => {
                if record.data.len() < 7 {
                    continue;
                }

                let sheet_offset =
                    u32::from_le_bytes([record.data[0], record.data[1], record.data[2], record.data[3]])
                        as usize;
                let Ok((name, _)) = strings::parse_biff_short_string(&record.data[6..], biff, encoding)
                else {
                    continue;
                };
                out.push(BoundSheetInfo {
                    name,
                    offset: sheet_offset,
                });
            }
            // EOF terminates the workbook global substream.
            0x000A => break,
            _ => {}
        }
    }

    Ok(out)
}

/// Workbook-global BIFF records needed for stable number format and date system import.
#[derive(Debug, Clone)]
pub(crate) struct BiffWorkbookGlobals {
    pub(crate) date_system: DateSystem,
    formats: HashMap<u16, String>,
    xfs: Vec<BiffXf>,
    pub(crate) warnings: Vec<String>,
}

impl Default for BiffWorkbookGlobals {
    fn default() -> Self {
        Self {
            date_system: DateSystem::Excel1900,
            formats: HashMap::new(),
            xfs: Vec::new(),
            warnings: Vec::new(),
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

        Some(format!(
            "{}{num_fmt_id}",
            formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX
        ))
    }

    pub(crate) fn xf_count(&self) -> usize {
        self.xfs.len()
    }
}

pub(crate) fn parse_biff_workbook_globals(
    workbook_stream: &[u8],
    biff: BiffVersion,
) -> Result<BiffWorkbookGlobals, String> {
    let encoding = strings::encoding_for_codepage(biff_codepage(workbook_stream));

    let mut out = BiffWorkbookGlobals::default();

    let mut saw_eof = false;
    let mut continuation_parse_failed = false;

    let allows_continuation: fn(u16) -> bool = match biff {
        BiffVersion::Biff5 => workbook_globals_allows_continuation_biff5,
        BiffVersion::Biff8 => workbook_globals_allows_continuation_biff8,
    };

    let iter = records::LogicalBiffRecordIter::new(workbook_stream, allows_continuation);

    for record in iter {
        let record = record?;
        let record_id = record.record_id;
        let data = record.data.as_ref();

        // BOF indicates the start of a new substream; the workbook globals contain
        // a single BOF at offset 0, so a second BOF means we're past the globals
        // section (even if the EOF record is missing).
        if record.offset != 0 && (record_id == 0x0809 || record_id == 0x0009) {
            saw_eof = true;
            break;
        }

        match record_id {
            // 1904 [MS-XLS 2.4.169]
            0x0022 => {
                if data.len() >= 2 {
                    let flag = u16::from_le_bytes([data[0], data[1]]);
                    if flag != 0 {
                        out.date_system = DateSystem::Excel1904;
                    }
                }
            }
            // FORMAT / FORMAT2 [MS-XLS 2.4.90]
            0x041E | 0x001E => match parse_biff_format_record_strict(&record, encoding) {
                Ok((num_fmt_id, code)) => {
                    out.formats.insert(num_fmt_id, code);
                }
                Err(_) if record.is_continued() => {
                    continuation_parse_failed = true;
                    if let Some((num_fmt_id, code)) =
                        parse_biff_format_record_best_effort(&record, encoding)
                    {
                        out.formats.insert(num_fmt_id, code);
                    }
                }
                Err(_) => {}
            },
            // XF [MS-XLS 2.4.353]
            0x00E0 => {
                if let Ok(xf) = parse_biff_xf_record(data) {
                    out.xfs.push(xf);
                }
            }
            // EOF terminates the workbook global substream.
            0x000A => {
                saw_eof = true;
                break;
            }
            _ => {}
        }
    }

    if continuation_parse_failed {
        out.warnings.push(
            "failed to parse one or more continued BIFF FORMAT records; number format codes may be truncated"
                .to_string(),
        );
    }

    if !saw_eof {
        // Some `.xls` files in the wild appear to be truncated or missing the
        // workbook-global EOF record. Treat this as a warning and return any
        // partial data we managed to parse so importers can still recover number
        // formats/date system where possible.
        out.warnings
            .push("unexpected end of workbook globals stream (missing EOF)".to_string());
    }

    Ok(out)
}

fn workbook_globals_allows_continuation_biff5(record_id: u16) -> bool {
    // FORMAT2 [MS-XLS 2.4.90]
    record_id == 0x001E
}

fn workbook_globals_allows_continuation_biff8(record_id: u16) -> bool {
    // FORMAT [MS-XLS 2.4.88]
    record_id == 0x041E
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

fn parse_biff_format_record_strict(
    record: &records::LogicalBiffRecord<'_>,
    encoding: &'static Encoding,
) -> Result<(u16, String), String> {
    let record_id = record.record_id;
    let data = record.data.as_ref();
    if data.len() < 2 {
        return Err("FORMAT record too short".to_string());
    }

    let num_fmt_id = u16::from_le_bytes([data[0], data[1]]);
    let rest = &data[2..];

    let mut code = match record_id {
        // BIFF8 FORMAT uses `XLUnicodeString` (16-bit length) and may be split
        // across one or more `CONTINUE` records.
        0x041E => {
            if record.is_continued() {
                let fragments: Vec<&[u8]> = record.fragments().collect();
                strings::parse_biff8_unicode_string_continued(&fragments, 2, encoding)?
            } else {
                strings::parse_biff8_unicode_string(rest, encoding)?.0
            }
        }
        // BIFF5 FORMAT2 uses a short ANSI string (8-bit length).
        0x001E => strings::parse_biff5_short_string(rest, encoding)?.0,
        _ => return Err(format!("unexpected FORMAT record id 0x{record_id:04X}")),
    };

    // Excel stores some strings with embedded NUL bytes; follow BoundSheet parsing and strip them.
    code = code.replace('\0', "");
    Ok((num_fmt_id, code))
}

fn parse_biff_format_record_best_effort(
    record: &records::LogicalBiffRecord<'_>,
    encoding: &'static Encoding,
) -> Option<(u16, String)> {
    let first = record.first_fragment();
    if first.len() < 2 {
        return None;
    }
    let num_fmt_id = u16::from_le_bytes([first[0], first[1]]);
    let rest = first.get(2..).unwrap_or_default();

    let mut code = match record.record_id {
        0x041E => strings::parse_biff8_unicode_string_best_effort(rest, encoding)?,
        0x001E => strings::parse_biff5_short_string_best_effort(rest, encoding)?,
        _ => return None,
    };
    code = code.replace('\0', "");
    Some((num_fmt_id, code))
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
    fn decodes_boundsheet_names_using_codepage() {
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(0x0042, &1251u16.to_le_bytes());

        // BoundSheet8 with a compressed 8-bit name (fHighByte=0).
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(0x80); // "Ђ" in cp1251
        let r_bs = record(0x0085, &bs_payload);

        let stream = [r_codepage, r_bs, record(0x000A, &[])].concat();
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "Ђ".to_string(),
                offset: 0x1234
            }]
        );
    }

    #[test]
    fn boundsheet_scan_stops_at_next_bof_without_eof() {
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(0x0042, &1251u16.to_le_bytes());

        // BoundSheet8 with a compressed 8-bit name (fHighByte=0).
        let mut bs_payload = Vec::new();
        bs_payload.extend_from_slice(&0x1234u32.to_le_bytes()); // sheet offset
        bs_payload.extend_from_slice(&[0, 0]); // visibility/type
        bs_payload.push(1); // cch
        bs_payload.push(0); // flags (compressed)
        bs_payload.push(0x80); // "Ђ" in cp1251
        let r_bs = record(0x0085, &bs_payload);

        // BOF for the next substream (worksheet).
        let r_sheet_bof = record(0x0809, &[0u8; 16]);

        // No EOF record; should still stop at the worksheet BOF.
        let stream = [r_codepage, r_bs, r_sheet_bof].concat();
        let sheets = parse_biff_bound_sheets(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(
            sheets,
            vec![BoundSheetInfo {
                name: "Ђ".to_string(),
                offset: 0x1234
            }]
        );
    }

    #[test]
    fn globals_scan_stops_at_next_bof_without_eof() {
        let r_bof_globals = record(0x0809, &[0u8; 16]);
        // CODEPAGE=1251 (Windows Cyrillic).
        let r_codepage = record(0x0042, &1251u16.to_le_bytes());

        // FORMAT id=200, code = byte 0x80 in cp1251 => "Ђ".
        let mut fmt_payload = Vec::new();
        fmt_payload.extend_from_slice(&200u16.to_le_bytes());
        fmt_payload.extend_from_slice(&1u16.to_le_bytes()); // cch
        fmt_payload.push(0); // flags (compressed)
        fmt_payload.push(0x80); // "Ђ" in cp1251
        let r_fmt = record(0x041E, &fmt_payload);

        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&200u16.to_le_bytes());
        let r_xf = record(0x00E0, &xf_payload);

        // BOF for the next substream (worksheet).
        let r_sheet_bof = record(0x0809, &[0u8; 16]);

        // A 1904 record and another CODEPAGE after the worksheet BOF should be ignored.
        let r_1904_after = record(0x0022, &[1, 0]);
        let r_codepage_after = record(0x0042, &1252u16.to_le_bytes());

        // No EOF for globals; parser should stop at the worksheet BOF.
        let stream = [
            r_bof_globals,
            r_codepage,
            r_fmt,
            r_xf,
            r_sheet_bof,
            r_1904_after,
            r_codepage_after,
        ]
        .concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1900);
        assert_eq!(globals.xf_count(), 1);
        assert_eq!(globals.resolve_number_format_code(0).as_deref(), Some("Ђ"));
    }

    #[test]
    fn globals_missing_eof_returns_partial_with_warning() {
        let r_bof_globals = record(0x0809, &[0u8; 16]);
        let r_1904 = record(0x0022, &[1, 0]);

        let mut xf_payload = vec![0u8; 20];
        xf_payload[2..4].copy_from_slice(&14u16.to_le_bytes()); // built-in date format
        let r_xf = record(0x00E0, &xf_payload);

        // No EOF record and no subsequent BOF; parser should return partial globals with a warning.
        let stream = [r_bof_globals, r_1904, r_xf].concat();
        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff8).expect("parse");
        assert_eq!(globals.date_system, DateSystem::Excel1904);
        assert_eq!(globals.xf_count(), 1);
        assert!(
            globals.warnings.iter().any(|w| w.contains("missing EOF")),
            "expected missing-EOF warning, got {:?}",
            globals.warnings
        );
        assert_eq!(
            globals.resolve_number_format_code(0).as_deref(),
            Some("m/d/yyyy")
        );
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

        let stream = [r_fmt, record(0x00E0, &xf_payload), record(0x000A, &[])].concat();

        let globals = parse_biff_workbook_globals(&stream, BiffVersion::Biff5).expect("parse");
        assert_eq!(globals.resolve_number_format_code(0).as_deref(), Some("0.00"));
    }
}
